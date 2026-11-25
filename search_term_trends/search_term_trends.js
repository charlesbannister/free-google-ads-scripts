/**
 * Search Term Trends Report
 * A Search Term Trends Report Google Ads Script
 * @author Charles Bannister
 * @version 1.1.0
 * More scripts at https://shabba.io
**/

// Template: https://docs.google.com/spreadsheets/d/1RQn6LN1H8shjKbZqbGpgmxPVjwWe-y4Pwf7ujIbWnM4
// Template: https://docs.google.com/spreadsheets/d/1RQn6LN1H8shjKbZqbGpgmxPVjwWe-y4Pwf7ujIbWnM4
// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1RQn6LN1H8shjKbZqbGpgmxPVjwWe-y4Pwf7ujIbWnM4/copy
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";


//Set to true or false, without quotes
//If true, an email will be sent to support if there's an error
//this will help us find and fix bugs
const SEND_ERROR_EMAIL_TO_SHABBA = false;

const SCRIPT_NAME = 'Search Term Trends Report';
const SHABBA_SCRIPT_ID = 11;


let INCLUDE_INTENT = false;

let LOCAL = false;
if (typeof AdsApp === 'undefined') {
  LOCAL = true;
}

let REPORT_DIMENSIONS = [
  'Search Term',
  'Campaigns',
];
if (INCLUDE_INTENT) {
  REPORT_DIMENSIONS.push('Intent');
}

const METRICS = {
  'clicks': 'Clicks per day',
  'impressions': 'Impressions per day',
  'conversions': 'Conversions per day',
  'conversionsValue': 'Conversion Value per day',
  'cost': 'Cost per day',
  'cos': 'COS%',
}

const LAST_UPDATED_RANGE_NAME = 'last_updated_timestamp';

const START_ROW = 17;
const START_COLUMN = 1;

let SPREADSHEET;
if (!LOCAL) {
  SPREADSHEET = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
}

const LOOKBACK_WINDOW_KEYS = ['date_1_days', 'date_2_days', 'date_3_days', 'date_4_days', 'date_5_days', 'date_6_days']


function main() {
  if (LOCAL) {
    runTopLevelLogic();
    return;
  }
  try {
    runTopLevelLogic();
  } catch (error) {
    sendErrorEmailToAdmin(error);
    throw new Error(error.stack);
  }
}

function runTopLevelLogic() {
  const sheetNames = getSheetNames();
  for (let sheetName of sheetNames) {
    if (sheetName === 'Notes') {
      continue;
    }
    runSheet(sheetName);
  }
}

function runSheet(sheetName) {
  const settings = new GetSettings(sheetName).getSettings();
  console.log('settings', JSON.stringify(settings));
  sheetSetup(sheetName);
  let searchTermData = {};
  updateSearchTermData(settings, sheetName, searchTermData);
  let sheetData = new GetSheetArray(searchTermData, settings).getSheetArray();
  console.log('Sheet Data: ', JSON.stringify(sheetData));
  writeSheetData(sheetData, sheetName);
  finalSheetFormatting(sheetName, settings['sort_by_column_number']);
  updateLastUpdatedTimestamp(sheetName);
}

function updateLastUpdatedTimestamp(sheetName) {
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);
  const lastUpdatedRange = reportSheet.getRange(LAST_UPDATED_RANGE_NAME);
  const now = new Date();
  lastUpdatedRange.setValue(now);
}

function updateSearchTermData(settings, sheetName, searchTermData) {
  for (let setting in settings) {
    if (!setting.includes('date_') || !setting.includes('_days')) {
      continue;
    }
    const lookbackDays = settings[setting];
    updateSearchTermDataForWindow(settings, lookbackDays, sheetName, searchTermData);
  }
}

class GetSettings {
  constructor(sheetName) {
    this.sheetName = sheetName;
    this.rangeNames = [
      'search_term_contains',
      'search_term_not_contains',
      'campaign_contains',
      'campaign_not_contains',
      'impressions_filter',
      'clicks_filter',
      'date_1_days',
      'date_1_name',
      'date_2_days',
      'date_2_name',
      'date_3_days',
      'date_3_name',
      'date_4_days',
      'date_4_name',
      'date_5_days',
      'date_5_name',
      'date_6_days',
      'date_6_name',
      'sort_by_column_number',
    ];
  }

  getSettings() {
    let settings = {};
    if (LOCAL) {
      //create dummy settings
      for (let rangeName of this.rangeNames) {
        settings[rangeName] = 'test';
      }
      //add some real dates for testing
      settings['date_1_days'] = 365;
      settings['date_1_name'] = 'Year';
      settings['date_6_days'] = 0;
      settings['date_6_name'] = 'Today';

      return settings;
    }
    const sheet = SPREADSHEET.getSheetByName(this.sheetName);

    for (let rangeName of this.rangeNames) {
      settings[rangeName] = sheet.getRange(rangeName).getValue();
    }
    return settings;
  }
}

class SheetFilter {

  constructor(sheetName) {
    this.sheetName = sheetName;
  }

  removeFilter() {
    if (LOCAL) {
      return;
    }
    const filterRange = this.getFilterRange();
    if (filterRange.getFilter()) {
      filterRange.getFilter().remove();
    }
  }

  createFilter() {
    if (LOCAL) {
      return;
    }
    const filterRange = this.getFilterRange();
    if (!filterRange.getFilter()) {
      filterRange.createFilter();
    }
  }

  getFilterRange() {
    const reportSheet = SPREADSHEET.getSheetByName(this.sheetName);
    const numRows = Math.max(reportSheet.getLastRow() - START_ROW, 1);
    const numColumns = Math.max(reportSheet.getLastColumn() - START_COLUMN + 1, 1);
    const filterRange = reportSheet.getRange(START_ROW + 1, START_COLUMN, numRows, numColumns);
    return filterRange;
  }
}

function finalSheetFormatting(sheetName, sortByColumnNumber) {
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);
  sortSheetValues(reportSheet, sortByColumnNumber);
  new SheetFilter(sheetName).createFilter();
}

function sortSheetValues(reportSheet, sortByColumnNumber) {
  if (LOCAL) {
    return;
  }
  const numRows = reportSheet.getLastRow() - START_ROW - 1;
  const numColumns = reportSheet.getLastColumn() - START_COLUMN + 1;
  const column = Number(sortByColumnNumber);
  if (numRows < 1 || !Number.isInteger(column) || column < START_COLUMN || column >= START_COLUMN + numColumns) {
    console.log('Not sorting: no data rows or sort_by_column_number is not a column of the report:', sortByColumnNumber);
    return;
  }
  reportSheet.getRange(START_ROW + 2, START_COLUMN, numRows, numColumns).sort({ column: column, ascending: false });
}

function updateSearchTermDataForWindow(settings, lookbackDays, sheetName, searchTermData) {
  let rawWindowSearchTermData = new GetSearchTermData(settings, lookbackDays, sheetName).getData();
  let windowSearchTermData = new GetReportData(rawWindowSearchTermData, lookbackDays).getData();
  searchTermData[String(lookbackDays)] = windowSearchTermData;
}

function writeSheetData(sheetData, sheetName) {
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);
  const lastRowNeeded = START_ROW + sheetData.length - 1;
  if (reportSheet.getMaxRows() < lastRowNeeded) {
    reportSheet.insertRowsAfter(reportSheet.getMaxRows(), lastRowNeeded - reportSheet.getMaxRows());
  }
  const lastColumnNeeded = START_COLUMN + sheetData[0].length - 1;
  if (reportSheet.getMaxColumns() < lastColumnNeeded) {
    reportSheet.insertColumnsAfter(reportSheet.getMaxColumns(), lastColumnNeeded - reportSheet.getMaxColumns());
  }
  reportSheet.getRange(START_ROW, START_COLUMN, sheetData.length, sheetData[0].length).setValues(sheetData);
}

function sheetSetup(sheetName) {
  console.log('setting up sheet:', sheetName);
  new SheetFilter(sheetName).removeFilter();
  clearSheetData(sheetName);
}
function clearSheetData(sheetName) {
  console.log('clearing sheet data');
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);

  if (reportSheet.getLastRow() < START_ROW + 1) {
    return;
  }
  reportSheet.getRange(START_ROW, START_COLUMN, reportSheet.getLastRow() - START_ROW + 1, reportSheet.getLastColumn() - START_COLUMN + 1).clear();
}

class GetSheetArray {

  constructor(searchTermData, settings) {
    this.searchTermData = searchTermData;
    this.settings = settings;
  }

  getSheetArray() {
    let reportData = this.searchTermData;
    let sheetArray = [this.getMetricsHeader()];
    sheetArray.push(this.getSheetArrayHeader());

    const firstWindowReportData = reportData[this.settings['date_1_days']];
    for (let searchTermText in firstWindowReportData) {
      let arrayRow = [];
      arrayRow.push(searchTermText);
      arrayRow.push(firstWindowReportData[searchTermText].campaigns);
      if (INCLUDE_INTENT) {
        arrayRow.push(firstWindowReportData[searchTermText].highestCampaignIntent);
      }

      for (let metric in METRICS) {
        for (let lookbackKey of LOOKBACK_WINDOW_KEYS) {
          let lookbackDays = this.settings[lookbackKey];
          if (searchTermText in reportData[lookbackDays] === false) {
            arrayRow.push(0);
            continue;
          }
          arrayRow.push(reportData[lookbackDays][searchTermText][metric]);
        }
      }
      sheetArray.push(arrayRow);
    }
    return sheetArray;

  }

  getMetricsHeader() {
    let header = [];
    for (let dimension of REPORT_DIMENSIONS) {
      header.push('');
    }
    for (let metric in METRICS) {
      header.push(METRICS[metric]);
      for (let i = 0; i < LOOKBACK_WINDOW_KEYS.length - 1; i++) {
        header.push('');
      }
    }
    return header;
  }

  getSheetArrayHeader() {
    let header = [];
    for (let dimension of REPORT_DIMENSIONS) {
      header.push(dimension);
    }
    for (let metric in METRICS) {
      for (let lookbackWindowKey of LOOKBACK_WINDOW_KEYS) {
        let lookbackWindowName = this.settings[lookbackWindowKey.replace('days', 'name')];
        header.push(lookbackWindowName);
      }
    }
    return header;

  }

}

class GetReportData {



  constructor(searchTermData, lookbackDays) {
    this.searchTermData = searchTermData;
    this.lookbackDays = lookbackDays;
  }


  getData() {
    let data = {};
    for (let searchTerm of this.searchTermData) {
      let searchTermText = searchTerm['search_term_view.search_term'];
      data[searchTermText] = data[searchTermText] || {};
      searchTerm['metrics.clicks'] = parseInt(searchTerm['metrics.clicks']);
      data[searchTermText].clicks = data[searchTermText].clicks + searchTerm['metrics.clicks'] || searchTerm['metrics.clicks'];
      searchTerm['metrics.impressions'] = parseInt(searchTerm['metrics.impressions']);
      // Report values can be strings: without this, "1.5" + "2" joins as "1.52"
      searchTerm['metrics.conversions'] = parseFloat(searchTerm['metrics.conversions']) || 0;
      searchTerm['metrics.conversions_value'] = parseFloat(searchTerm['metrics.conversions_value']) || 0;
      data[searchTermText].impressions = data[searchTermText].impressions + searchTerm['metrics.impressions'] || searchTerm['metrics.impressions'];
      data[searchTermText].conversions = data[searchTermText].conversions + searchTerm['metrics.conversions'] || searchTerm['metrics.conversions'];
      data[searchTermText].conversionsValue = data[searchTermText].conversionsValue + searchTerm['metrics.conversions_value'] || searchTerm['metrics.conversions_value'];
      let cost = searchTerm['metrics.cost_micros'] / 1000000;
      data[searchTermText].cost = data[searchTermText].cost + cost || cost;
      data[searchTermText].cos = calculateCostOfSalePercentage(data[searchTermText].cost, data[searchTermText].conversionsValue);
      // A search_term_view row is one term in one ad group, so list each campaign once
      if (typeof data[searchTermText].campaigns === 'undefined') {
        data[searchTermText].campaigns = searchTerm["campaign.name"];
      } else if (data[searchTermText].campaigns.split(", ").indexOf(searchTerm['campaign.name']) === -1) {
        data[searchTermText].campaigns = data[searchTermText].campaigns + ", " + searchTerm['campaign.name'];
      }
      if (INCLUDE_INTENT) {
        data[searchTermText].highestCampaignIntent = this.getHighestCampaignIntent(data[searchTermText].campaigns);
      }
    }

    this.averageValues(data);

    //round the values
    for (let searchTermText in data) {
      data[searchTermText].clicks = Math.round(data[searchTermText].clicks * 100) / 100;
      data[searchTermText].impressions = Math.round(data[searchTermText].impressions * 100) / 100;
      data[searchTermText].conversions = Math.round(data[searchTermText].conversions * 100) / 100;
      data[searchTermText].conversionsValue = Math.round(data[searchTermText].conversionsValue * 100) / 100;
      data[searchTermText].cost = Math.round(data[searchTermText].cost * 100) / 100;
    }
    return data;
  }

  averageValues(data) {
    // 0 (today) or a blank/invalid window: don't divide (it would give Infinity or NaN)
    if (!(Number(this.lookbackDays) > 0)) {
      return data;
    }

    //average the values by day
    for (let searchTermText in data) {
      data[searchTermText].clicks = data[searchTermText].clicks / this.lookbackDays;
      data[searchTermText].impressions = data[searchTermText].impressions / this.lookbackDays;
      data[searchTermText].conversions = data[searchTermText].conversions / this.lookbackDays;
      data[searchTermText].conversionsValue = data[searchTermText].conversionsValue / this.lookbackDays;
      data[searchTermText].cost = data[searchTermText].cost / this.lookbackDays;
    }

  }


  /**
   * Return High, Medium or Low based on which campaigns the searchTerm falls under
   * @param {string} campaigns
   */
  getHighestCampaignIntent(campaigns) {
    if (campaigns.includes('High Intent')) {
      return 'High';
    }
    if (campaigns.includes('Medium Intent')) {
      return 'Medium';
    }
    if (campaigns.includes('Low Intent')) {
      return 'Low';
    }
    return '--';
  }

}



class GetSearchTermData {

  constructor(settings, lookbackDays, sheetName) {
    this.settings = settings;
    this.lookbackDays = lookbackDays;
    this.sheetName = sheetName;
  }

  getData() {
    if (LOCAL) {
      //get the query here just to review/test it
      const query = this.getApiQuery();
      return this.getMockData();

    }
    return this.getApiData();
  }

  getApiData() {
    const query = this.getApiQuery();
    console.log('api query:', query);
    const report = AdsApp.report(query);
    const rows = report.rows();
    let data = [];
    if (!rows.hasNext()) {
      console.error('No data found for query:', query);
    }
    while (rows.hasNext()) {
      data.push(rows.next());
    }
    return data;
  }

  getApiQuery() {
    let query = `SELECT search_term_view.search_term, metrics.clicks, 
      metrics.conversions, metrics.conversions_value, metrics.cost_micros, metrics.impressions `;
    query += ` , campaign.name `;
    query += `from search_term_view `;
    query += ` where metrics.impressions > 0 `;
    query += this.getWhereStringFromSettings();
    query += this.getDateRangeString();
    return query;
  }

  getWhereStringFromSettings() {
    let whereString = '';
    if (this.settings['search_term_contains'] !== '') {
      whereString += ` and search_term_view.search_term like "%${likeValue(this.settings['search_term_contains'])}%" `;
    }
    if (this.settings['search_term_not_contains'] !== '') {
      whereString += ` and search_term_view.search_term not like "%${likeValue(this.settings['search_term_not_contains'])}%" `;
    }
    if (this.settings['campaign_contains'] !== '') {
      whereString += ` and campaign.name like "%${likeValue(this.settings['campaign_contains'])}%" `;
    }
    if (this.settings['campaign_not_contains'] !== '') {
      whereString += ` and campaign.name not like "%${likeValue(this.settings['campaign_not_contains'])}%" `;
    }
    if (isNumberSetting(this.settings['impressions_filter'])) {
      whereString += ` and metrics.impressions > ${Number(this.settings['impressions_filter'])}`;
    } else if (this.settings['impressions_filter'] !== '') {
      console.warn('Ignoring impressions_filter, it is not a number:', this.settings['impressions_filter']);
    }
    if (isNumberSetting(this.settings['clicks_filter'])) {
      whereString += ` and metrics.clicks > ${Number(this.settings['clicks_filter'])}`;
    } else if (this.settings['clicks_filter'] !== '') {
      console.warn('Ignoring clicks_filter, it is not a number:', this.settings['clicks_filter']);
    }
    return whereString;
  }

  getDateRangeString() {
    let endDate = new Date();
    if (this.lookbackDays > 0) {
      endDate.setDate(endDate.getDate() - 1);
    }
    let startDate = new Date();
    startDate.setDate(startDate.getDate() - this.lookbackDays);
    return ` and segments.date >= '${this.dateRangeStringFromDate(startDate)}' and segments.date <= '${this.dateRangeStringFromDate(endDate)}' `;
  }

  dateRangeStringFromDate(date) {
    //format the date as YYYY-MM-DD
    let year = date.getFullYear();
    let month = date.getMonth() + 1 < 10 ? `0${date.getMonth() + 1}` : date.getMonth() + 1;
    let day = date.getDate() < 10 ? `0${date.getDate()}` : date.getDate();
    return `${year}-${month}-${day}`;
  }

  getMockData() {
    return [
      {
        "search_term_view.search_term": "ken doll",
        "metrics.clicks": "7",
        "metrics.impressions": "76",
        "metrics.conversions": 0,
        "metrics.conversions_value": 0,
        "metrics.cost_micros": 10000000,
        "campaign.name": "Shopping | Men's Jackets | Low Intent | CPC | UK"
      },
      {
        "search_term_view.search_term": "ken doll",
        "metrics.clicks": "85",
        "metrics.impressions": "762",
        "metrics.conversions": 4.663353,
        "metrics.conversions_value": 232.863323,
        "metrics.cost_micros": 100000000,
        "campaign.name": "Shopping | Men's Jackets | Medium Intent | tROAS 330% | UK"
      },
    ];
  }
}

/**
 * Makes a sheet value safe inside a double-quoted GAQL string:
 * escapes \ and " so they cannot break the query. Wildcards (%, _) work as before.
 */
function likeValue(value) {
  return String(value)
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"');
}

function isNumberSetting(value) {
  return value !== '' && value !== null && String(value).trim() !== '' && isFinite(Number(value));
}

function calculateCostOfSalePercentage(cost, revenue) {
  if (revenue === 0 || cost === 0) {
    return 0;
  }
  let costOfSale = cost / revenue;
  let costOfSalePercentage = costOfSale * 100;
  let costOfSalePercentageRounded = Math.round(costOfSalePercentage * 100) / 100;
  return costOfSalePercentageRounded + '%';
}


function log() {
  console.log(...arguments);
}

if (LOCAL) {
  main();
}

function getSheetNames() {
  if (LOCAL) {
    return ['UK'];
  }
  let sheetNames = [];
  for (let sheet of SPREADSHEET.getSheets()) {
    sheetNames.push(sheet.getName());
  }
  return sheetNames;
}

function sendErrorEmailToAdmin(error) {
  if (LOCAL) {
    return;
  }
  if (typeof SEND_ERROR_EMAIL_TO_SHABBA === 'undefined') {
    return;
  }
  if (!SEND_ERROR_EMAIL_TO_SHABBA) {
    return;
  }
  const adminEmail = "charles@shabba.io";
  const subject = `Shabba.io ${SCRIPT_NAME} Script Error`;
  let body = `A user got an error.\n`;
  body += `Shabba Script ID: ${SHABBA_SCRIPT_ID}\n`;
  if (typeof loser !== 'undefined') {
    body += `User ID: ${loser}\n`
  } else {
    body += `User ID: Unknown\n`
  }
  body += `They got the following error: \n`
  body += error;
  body += `Here is the stack: \n`
  body += error.stack;
  MailApp.sendEmail(adminEmail, subject, body);
}