/**
* Google Shopping Trends Report
 * @author Charles Bannister
 * More scripts at https://shabba.io
 * @version 1.1.0
**/

// Template: https://docs.google.com/spreadsheets/d/1oxPn_cd-2lm5nKUof7GMgDDgBrMjfm2ZQOB3kNmgkWU/
// Template: https://docs.google.com/spreadsheets/d/1oxPn_cd-2lm5nKUof7GMgDDgBrMjfm2ZQOB3kNmgkWU
// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1oxPn_cd-2lm5nKUof7GMgDDgBrMjfm2ZQOB3kNmgkWU/copy
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";


// This setting will pull the Intent should campaigns
// be named with "High Intent", "Medium Intent" or "Low Intent"
const INCLUDE_INTENT = false;

const VERSION = '1.1.0';
const UPDATED_AT = 'April 2024'

//Set to true or false, without quotes
//If true, an email will be sent to support if there's an error
//then, if necessary, we'll look into the problem and get back to you
const SEND_ERROR_EMAIL_TO_SHABBA = false;

const SCRIPT_NAME = 'Barry';
const SHABBA_SCRIPT_ID = 10;

let LOCAL = false;
if (typeof AdsApp === 'undefined') {
  LOCAL = true;
}

const NOTES_SHEET_NAME = 'Notes';

let REPORT_DIMENSIONS = [
  'Item ID',
  'Title',
  'Campaigns',
];
if (typeof INCLUDE_INTENT !== 'undefined' && INCLUDE_INTENT) {
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

const LAST_UPDATED_RANGE_NAME = 'B1';

const START_ROW = 17;
const START_COLUMN = 1;

let SPREADSHEET;
if (!LOCAL) {
  SPREADSHEET = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
  SPREADSHEET.getSheetByName(NOTES_SHEET_NAME).getRange('B2').setValue(VERSION);
  SPREADSHEET.getSheetByName(NOTES_SHEET_NAME).getRange('B4').setValue(UPDATED_AT);
}

const LOOKBACK_WINDOW_KEYS = ['date_1_days', 'date_2_days', 'date_3_days', 'date_4_days', 'date_5_days', 'date_6_days']


function main() {
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
    if (sheetName === NOTES_SHEET_NAME) {
      continue;
    }
    runSheet(sheetName);
  }
}

function runSheet(sheetName) {
  const settings = new GetSettings(sheetName).getSettings();
  console.log('settings', JSON.stringify(settings));
  sheetSetup(sheetName);
  let productData = {};
  updateProductData(settings, sheetName, productData);
  let sheetData = new GetSheetArray(productData, settings).getSheetArray();
  writeSheetData(sheetData, sheetName);
  finalSheetFormatting(sheetName, settings['sort_by_column_number']);
  updateLastUpdatedTimestamp(sheetName);
  console.log(`Finished updating sheet: ${sheetName}`);
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

function updateProductData(settings, sheetName, productData) {
  for (let setting in settings) {
    if (!setting.includes('date_') || !setting.includes('_days')) {
      continue;
    }
    const lookbackDays = settings[setting];
    updateProductDataForWindow(settings, lookbackDays, sheetName, productData);
  }
}

class GetSettings {
  constructor(sheetName) {
    this.sheetName = sheetName;
    this.rangeByName = {
      'item_id_contains': 'A7',
      'item_id_not_contains': 'B7',
      'product_title_contains': 'A10',
      'product_title_not_contains': 'B10',
      'campaign_contains': 'A13',
      'campaign_not_contains': 'B13',
      'date_1_days': 'D7',
      'date_1_name': 'E7',
      'date_2_days': 'D8',
      'date_2_name': 'E8',
      'date_3_days': 'D9',
      'date_3_name': 'E9',
      'date_4_days': 'D10',
      'date_4_name': 'E10',
      'date_5_days': 'D11',
      'date_5_name': 'E11',
      'date_6_days': 'D12',
      'date_6_name': 'E12',
      'sort_by_column_number': 'G7',
    };
  }

  getSettings() {
    console.log(`Getting settings for sheet: ${this.sheetName}`);
    let settings = {};
    if (LOCAL) {
      //create dummy settings
      for (let rangeName in this.rangeByName) {
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

    for (let rangeName in this.rangeByName) {
      const cellReference = this.rangeByName[rangeName];
      settings[rangeName] = sheet.getRange(cellReference).getValue();
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
    if (!filterRange) {
      return;
    }
    if (filterRange.getFilter()) {
      filterRange.getFilter().remove();
    }
  }

  createFilter() {
    if (LOCAL) {
      return;
    }
    const filterRange = this.getFilterRange();
    if (!filterRange) {
      return;
    }
    if (!filterRange.getFilter()) {
      filterRange.createFilter();
    }
  }

  getFilterRange() {
    const reportSheet = SPREADSHEET.getSheetByName(this.sheetName);
    const rangeValues = [START_ROW + 1, START_COLUMN, reportSheet.getLastRow(), reportSheet.getLastColumn()];
    const filterRange = reportSheet.getRange(...rangeValues);
    return filterRange;
  }
}

function finalSheetFormatting(sheetName, sortByColumnNumber) {
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);
  new SheetFilter(sheetName).createFilter();
  sortSheetValues(reportSheet, sortByColumnNumber);
}

function sortSheetValues(reportSheet, sortByColumnNumber,) {
  if (LOCAL) {
    return;
  }
  reportSheet.getRange(START_ROW + 2, START_COLUMN, getLastRow(reportSheet), reportSheet.getLastColumn()).sort({ column: sortByColumnNumber, ascending: false });
}

/**
 * The sheet.getLastRow() method doesn't work as expected
 * so I'm using this for now
 * @param {object} sheet 
 * @returns {number}
 */
function getLastRow(sheet) {
  const data = sheet.getDataRange().getValues();
  let lastRow = START_ROW + 1;
  for (let row of data) {
    if (row[0] === '') {
      break;
    }
    lastRow++;
  }
  return lastRow;
}

function updateProductDataForWindow(settings, lookbackDays, sheetName, productData) {
  let rawWindowProductData = new GetProductData(settings, lookbackDays, sheetName).getData();
  let windowProductData = new GetReportData(rawWindowProductData, lookbackDays).getData();
  productData[String(lookbackDays)] = windowProductData;
}

function writeSheetData(sheetData, sheetName) {
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);
  reportSheet.getRange(START_ROW, START_COLUMN, sheetData.length, sheetData[0].length).setValues(sheetData);
}

function sheetSetup(sheetName) {
  new SheetFilter(sheetName).removeFilter();
  clearSheetData(sheetName);
}
function clearSheetData(sheetName) {
  if (LOCAL) {
    return;
  }
  const reportSheet = SPREADSHEET.getSheetByName(sheetName);

  if (reportSheet.getLastRow() < START_ROW + 1) {
    return;
  }
  reportSheet.getRange(START_ROW, START_COLUMN, reportSheet.getLastRow(), reportSheet.getLastColumn()).clear();
}

class GetSheetArray {

  constructor(productData, settings) {
    this.productData = productData;
    this.settings = settings;
  }

  getSheetArray() {
    let reportData = this.productData;
    let sheetArray = [this.getMetricsHeader()];
    sheetArray.push(this.getSheetArrayHeader());

    const firstWindowReportData = reportData[this.settings['date_1_days']];
    for (let itemId in firstWindowReportData) {
      let arrayRow = [];
      arrayRow.push(itemId);
      arrayRow.push(firstWindowReportData[itemId].title);
      arrayRow.push(firstWindowReportData[itemId].campaigns);
      if (typeof INCLUDE_INTENT !== 'undefined' && INCLUDE_INTENT) {
        arrayRow.push(firstWindowReportData[itemId].highestCampaignIntent);
      }

      for (let metric in METRICS) {
        for (let lookbackKey of LOOKBACK_WINDOW_KEYS) {
          let lookbackDays = this.settings[lookbackKey];
          if (itemId in reportData[lookbackDays] === false) {
            arrayRow.push(0);
            continue;
          }
          arrayRow.push(reportData[lookbackDays][itemId][metric]);
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



  constructor(productData, lookbackDays) {
    this.productData = productData;
    this.lookbackDays = lookbackDays;
  }


  getData() {
    let data = {};
    for (let product of this.productData) {
      let itemId = product['segments.product_item_id'];
      this._formatMetrics(product);
      data[itemId] = data[itemId] || {};
      data['total'] = data['total'] || {};
      data[itemId].title = product['segments.product_title'];
      data['total'].title = 'Total';
      this._updateItemIdStats(data, itemId, product);
      this._updateItemIdStats(data, 'total', product);
      data['total'].campaigns = '';
      data[itemId].campaigns = typeof data[itemId].campaigns === 'undefined' ? product["campaign.name"] : data[itemId].campaigns + ", " + product['campaign.name'];
      if (typeof INCLUDE_INTENT !== 'undefined' && INCLUDE_INTENT) {
        data[itemId].highestCampaignIntent = this.getHighestCampaignIntent(data[itemId].campaigns);
        data['total'].highestCampaignIntent = '';
      }
    }

    this.averageValues(data);

    //round the values
    for (let itemId in data) {
      data[itemId].clicks = Math.round(data[itemId].clicks * 100) / 100;
      data[itemId].impressions = Math.round(data[itemId].impressions * 100) / 100;
      data[itemId].conversions = Math.round(data[itemId].conversions * 100) / 100;
      data[itemId].conversionsValue = Math.round(data[itemId].conversionsValue * 100) / 100;
      data[itemId].cost = Math.round(data[itemId].cost * 100) / 100;
    }
    return data;
  }

  _formatMetrics(product) {
    product['metrics.clicks'] = parseInt(product['metrics.clicks']);
    product['metrics.impressions'] = parseInt(product['metrics.impressions']);
    product['metrics.conversions'] = parseFloat(product['metrics.conversions']);
    product['metrics.conversions_value'] = parseFloat(product['metrics.conversions_value']);
    product['cost'] = product['metrics.cost_micros'] / 1000000;
  }

  _updateItemIdStats(data, itemId, product) {
    data[itemId].clicks = data[itemId].clicks + product['metrics.clicks'] || product['metrics.clicks'];
    data[itemId].impressions = data[itemId].impressions + product['metrics.impressions'] || product['metrics.impressions'];
    data[itemId].conversions = data[itemId].conversions + product['metrics.conversions'] || product['metrics.conversions'];
    data[itemId].conversionsValue = data[itemId].conversionsValue + product['metrics.conversions_value'] || product['metrics.conversions_value'];
    data[itemId].cost = data[itemId].cost + product['cost'] || product['cost'];
    data[itemId].cos = calculateCostOfSalePercentage(data[itemId].cost, data[itemId].conversionsValue);
  }

  averageValues(data) {
    if (this.lookbackDays === 0) {
      return data;
    }

    //average the values by day
    for (let itemId in data) {
      data[itemId].clicks = data[itemId].clicks / this.lookbackDays;
      data[itemId].impressions = data[itemId].impressions / this.lookbackDays;
      data[itemId].conversions = data[itemId].conversions / this.lookbackDays;
      data[itemId].conversionsValue = data[itemId].conversionsValue / this.lookbackDays;
      data[itemId].cost = data[itemId].cost / this.lookbackDays;
    }

  }


  /**
   * Return High, Medium or Low based on which campaigns the product falls under
   * @param {string} campaigns
   */
  getHighestCampaignIntent(campaigns) {
    if (campaigns.includes('High Intent')) {
      return 'High';
    }
    if (campaigns.includes('Medium Intent')) {
      return 'Medium';
    }
    return 'Low';
  }

}



class GetProductData {

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
    let query = `SELECT segments.product_item_id, segments.product_title, metrics.clicks, 
      metrics.conversions, metrics.conversions_value, metrics.cost_micros, metrics.impressions `;
    query += ` , campaign.name `;
    query += `from shopping_performance_view `;
    query += ` where metrics.impressions > 0 `;
    query += this.getWhereStringFromSettings();
    query += this.getDateRangeString();
    return query;
  }

  getWhereStringFromSettings() {
    let whereString = '';
    if (this.settings['item_id_contains'] !== '') {
      whereString += ` and segments.product_item_id like "%${this.settings['item_id_contains']}%" `;
    }
    if (this.settings['item_id_not_contains'] !== '') {
      whereString += ` and segments.product_item_id not like "%${this.settings['item_id_not_contains']}%" `;
    }
    if (this.settings['product_title_contains'] !== '') {
      whereString += ` and segments.product_title like "%${this.settings['product_title_contains']}%" `;
    }
    if (this.settings['product_title_not_contains'] !== '') {
      whereString += ` and segments.product_title not like "%${this.settings['product_title_not_contains']}%" `;
    }
    if (this.settings['campaign_contains'] !== '') {
      whereString += ` and campaign.name like "%${this.settings['campaign_contains']}%" `;
    }
    if (this.settings['campaign_not_contains'] !== '') {
      whereString += ` and campaign.name not like "%${this.settings['campaign_not_contains']}%" `;
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
        "segments.product_item_id": "jack102a",
        "segments.product_title": "Men's Very Nice Red Jacket with 1,000 pockets for all your things",
        "metrics.clicks": "7",
        "metrics.impressions": "76",
        "metrics.conversions": 0,
        "metrics.conversions_value": 0,
        "metrics.cost_micros": 10000000,
        "campaign.name": "Shopping | Men's Jackets | Low Intent | CPC | UK"
      },
      {
        "segments.product_item_id": "jack102a",
        "segments.product_title": "Men's Very Nice Red Jacket with 1,000 pockets for all your things",
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