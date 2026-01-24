/*************************************************
* Product Performance Alerts & Reports
* Single account script (contact me for MCC support)
* This script cannot make changes to your account.
* @author Charles Bannister
* @link https://shabba.io/script/16
* @version: 1.0.0
***************************************************/

// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1hxdVWZ8LPZzZEfevniOvmwYhIdrtgK6cAb5QQWUQCv4/copy
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";




let INPUT_TAB_NAME = "Settings";

const SCRIPT_NAME = "Product Performance Alerts & Reports"

const VERSION = "1.0.0";

let NUMBER_OF_FILTERS = 6;

const DEBUG = false;
const LOCAL = false;

//No need to edit anything below this line
let h = new Helper();
let s = new Setting();

let CUSTOM_METRICS = {
  //metric, metric, operator (divide or multiply, whether high or low is good)
  Ctr: ["metrics.clicks", "metrics.impressions", "divide", "high"],
  Roas: ["metrics.conversions_value", "metrics.cost", "divide", "high"],
  Cos: ["metrics.cost", "metrics.conversions_value", "divide", "low"],
  Cpa: ["metrics.cost", "metrics.conversions", "divide", "low"],
  AverageCpc: ["metrics.cost", "metrics.clicks", "divide", "low"],
  ConversionRate: ["metrics.conversions", "metrics.clicks", "divide", "high"],
  Rpc: ["metrics.conversions_value", "metrics.clicks", "divide", "high"],
};

function runScript(SETTINGS) {
  log("Script Started");
  SETTINGS.CONTROL_SHEET =
    SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME);

  checkSettings(SETTINGS);

  processRowSettings(SETTINGS);

  addLogSheetInfo(SETTINGS);

  log(JSON.stringify(SETTINGS));

  const shouldEmail = addProductsToSheet(SETTINGS);

  SETTINGS.LOGS.push("The script ran successfully");
  updateControlSheet("", SETTINGS);

  if (shouldEmail && SETTINGS.SEND_EMAIL) {
    sendEmail(SETTINGS);
  }

  log("Finished");
}

/**
 * Retrieves the resource name of a label with the same name as the provided string.
 * @param {string} str - The name of the label to retrieve the resource name for.
 * @returns {string} - The resource name of the label, or a warning message if no label is found.
 */
function getLabelResourceNameFromString(str) {
  let labels = AdsApp.labels()
    .withCondition("Name = '" + str.trim() + "'")
    .get();
  if (labels.hasNext()) {
    let label = labels.next();
    return label.getResourceName();
  } else {
    console.warn("Error: The label with name '" + str + "' cannot be found");
  }
}

function getLabelTextFromResourceName(resourceNames) {
  if (!resourceNames) {
    return "";
  }
  let labels = AdsApp.labels()
    .withResourceNames(resourceNames)
    .get();
  let labelText = []
  if (!labels.hasNext()) {
    console.warn("Error: The label with resource names '" + resourceNames + "' cannot be found");
  }
  while (labels.hasNext()) {
    let label = labels.next();
    labelText.push(label.getName());
  }
  return labelText.join(",")
}

function addLogSheetInfo(SETTINGS) {
  SETTINGS.LOG_SHEET_URL = INPUT_SHEET_URL;

  let logSS = SpreadsheetApp.openByUrl(SETTINGS.LOG_SHEET_URL);
  SETTINGS.logSS = logSS;
}

function sendEmail(SETTINGS) {
  if (DEBUG) {
    return;
  }
  const prefix = SETTINGS.EMAIL_PREFIX ? SETTINGS.EMAIL_PREFIX + " " : "";
  let subject = prefix + AdsApp.currentAccount().getName() + " - " + SCRIPT_NAME + " script.";
  let message =
    "Hi,<br><br>The " +
    SCRIPT_NAME +
    " script has results for the " + SETTINGS.NAME + " rule.<br><br>";

  message +=
    "<br><br>Please follow the link below for more information:<br>" +
    SETTINGS.LOG_SHEET_URL;
  message += "<br><br>Here are the settings:<br><br>";

  message += "<ul>";
  for (let settingsKey in SETTINGS) {
    message += "<li>";
    message += settingsKey + ": " + SETTINGS[settingsKey];
    message += "</li>";
  }
  message += "</ul>";


  message += "<br><br>Thanks.";
  let emails = SETTINGS.EMAILS;

  for (let i in emails) {
    MailApp.sendEmail({
      to: emails[i],
      subject: subject,
      htmlBody: message,
    });
  }
}

function updateControlSheet(errorMessage, SETTINGS) {
  //remove the unsuccessful note we added
  SETTINGS.CONTROL_SHEET.getRange(
    SETTINGS.ROW_NUM,
    SETTINGS.LOGS_COLUMN,
    1,
    1
  ).setNote("");

  if (errorMessage != "") {
    SETTINGS.CONTROL_SHEET.getRange(
      SETTINGS.ROW_NUM,
      SETTINGS.LOGS_COLUMN,
      1,
      2
    ).setValues([[errorMessage, SETTINGS.NOW]]);
    SETTINGS.CONTROL_SHEET.getRange(
      SETTINGS.ROW_NUM,
      SETTINGS.LOGS_COLUMN,
      1,
      1
    ).setNote(
      "Note: Some rows running on account ID " +
      SETTINGS.ID +
      " may not have completed sucessfully. Please see their respective logs and 'Last Run' times."
    );
    if (!errorMessage.includes(" is 0, skipped bid update.")) {
      throw errorMessage;
    } else {
      SETTINGS.LOGS.push(errorMessage);
    }
  }

  //add final logs
  //stringify logs
  logString = h.stringifyLogs(SETTINGS.LOGS);
  //update control sheet
  let put = [logString, SETTINGS.NOW];
  SETTINGS.CONTROL_SHEET.getRange(
    SETTINGS.ROW_NUM,
    SETTINGS.LOGS_COLUMN,
    1,
    2
  ).setValues([put]);
  SETTINGS.CONTROL_SHEET.getRange(
    SETTINGS.ROW_NUM,
    SETTINGS.LOGS_COLUMN,
    1,
    1
  ).setNote(logString);
}

function swapLabelTextForIds(SETTINGS) {
  //swap label text for label ids
  for (let key in SETTINGS) {
    if (key.indexOf("FILTER") === -1 || key.indexOf("METRIC") === -1) {
      continue;
    }
    if (SETTINGS[key] !== "Labels") continue;
    log(key + " - " + SETTINGS[key]);
    let filter_value = SETTINGS[key.replace("METRIC", "VALUE")];
    let value_split = String(filter_value).split(",");
    for (let v in value_split) {
      value_split[v] = getLabelResourceNameFromString(value_split[v], SETTINGS);
    }
    SETTINGS[key.replace("METRIC", "VALUE")] = value_split.join(",");
  }
  return SETTINGS;
}

// log(getFilterWhereString(SETTINGS))

function getFilterWhereString(SETTINGS) {
  console.log("getFilterWhereString...");

  const filterMap = extractFilterMap(SETTINGS);
  const whereArray = buildWhereArray(filterMap);
  const where = whereArray.join(" ");

  return where;
}

function extractFilterMap(SETTINGS) {
  const filters = Object.keys(SETTINGS).filter(x => x.indexOf("FILTER") > -1);
  const filterMap = {};
  const filterParts = ["metric", "operator", "value"];
  const numberOfFilters = filters.length / filterParts.length;

  for (let i = 0; i < numberOfFilters; i++) {
    const filterName = "FILTER_" + (i + 1);
    //note clicks will be skipped here
    //it's already been added to the where string
    if (Object.keys(SETTINGS).indexOf(filterName + "_METRIC") === -1) {
      continue
    };
    if (SETTINGS[filterName + "_METRIC"] === "") {
      continue
    }
    filterMap[filterName] = {};
    for (const part of filterParts) {
      filterMap[filterName][part] = SETTINGS[filterName + "_" + part.toUpperCase()];
    }
  }

  console.log(`Note: FILTER_1 will be removed here if it's clicks - it gets applied to the GAQL query`)
  console.log(`filterMap: ${JSON.stringify(filterMap)}`);

  return filterMap;
}

function buildWhereArray(filterMap) {
  const whereArray = [];
  const filterParts = ["metric", "operator", "value"];
  for (const filter in filterMap) {

    if (filterMap[filter]['metric'] === '' || typeof filterMap[filter]['metric'] === 'undefined') {
      continue
    }
    if (filterMap[filter]['operator'] === "CONTAINS" || filterMap[filter]['operator'] === "CONTAINS_IGNORE_CASE") {
      filterMap[filter]['operator'] = 'LIKE'
      filterMap[filter]['value'] = `'%${filterMap[filter]['value'].replace("'", "").replace("'", "")}%'`
    }
    if (filterMap[filter]['operator'].indexOf("DOES_NOT_CONTAIN") > -1) {
      filterMap[filter]['operator'] = 'NOT LIKE'
      filterMap[filter]['value'] = `'%${filterMap[filter]['value'].replace("'", "").replace("'", "")}%'`
    }
    if (filterMap[filter]['operator'].indexOf("CONTAINS ") > -1) {
      filterMap[filter]['value'] = filterMap[filter]['value'].replace("'", "").replace("'", "").split(",").map((x) => { return `'${x}'` }).join(",")
      filterMap[filter]['value'] = `(${filterMap[filter]['value']})`
    }
    log(`JSON.stringify(filterMap[filter]): ${JSON.stringify(filterMap[filter])}`)
    if (Object.keys(CUSTOM_METRICS).indexOf(filterMap[filter]["metric"]) > -1) {
      continue;
    }
    if (String(filterMap[filter]["value"]).toLowerCase().indexOf("avg") > -1) {
      continue;
    }
    const partsArray = [];
    for (const part of filterParts) {
      partsArray.push(filterMap[filter][part]);
    }
    whereArray.push("AND " + partsArray.join(" "));
  }

  return whereArray;
}

function isEmptyObject(obj) {
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      return false;
    }
  }
  return true;
}

function getQueryWhereString(SETTINGS) {

  let where = "where metrics.impressions > 0 ";
  if (SETTINGS.INCLUDE_CAMPAIGN) {
    where += "and campaign.status in  (ENABLED) ";
  }
  if (SETTINGS.INCLUDE_AD_GROUP) {
    where += "and ad_group.status in  (ENABLED) ";
  }

  if (SETTINGS.FILTER_1_METRIC === "metrics.clicks") {
    where +=
      " and metrics.clicks " +
      SETTINGS.FILTER_1_OPERATOR +
      " " +
      SETTINGS.FILTER_1_VALUE;

    //preent this being added twice
    SETTINGS.FILTER_1_METRIC = "";
    SETTINGS.FILTER_1_OPERATOR = "";
    SETTINGS.FILTER_1_VALUE = "";
  }

  let whereArray = [];

  for (let i in SETTINGS.CAMPAIGN_NAME_CONTAINS) {
    whereArray.push(
      " and campaign.name LIKE '%" +
      SETTINGS.CAMPAIGN_NAME_CONTAINS[i].trim() +
      "%'"
    );
  }

  for (let i in SETTINGS.CAMPAIGN_NAME_NOT_CONTAINS) {
    whereArray.push(
      " and campaign.name NOT LIKE '%" +
      SETTINGS.CAMPAIGN_NAME_NOT_CONTAINS[i].trim() +
      "%'"
    );
  }

  for (let i in SETTINGS.ITEM_ID_CONTAINS) {
    whereArray.push(
      " and segments.product_item_id LIKE '%" +
      SETTINGS.ITEM_ID_CONTAINS[i].trim() +
      "%'"
    );
  }

  for (let i in SETTINGS.ITEM_ID_NOT_CONTAINS) {
    whereArray.push(
      " and segments.product_item_id NOT LIKE '%" +
      SETTINGS.ITEM_ID_NOT_CONTAINS[i].trim() +
      "%'"
    );
  }

  for (let i in SETTINGS.PRODUCT_TITLE_CONTAINS) {
    whereArray.push(
      " and segments.product_title LIKE '%" +
      SETTINGS.PRODUCT_TITLE_CONTAINS[i].trim() +
      "%'"
    );
  }

  for (let i in SETTINGS.PRODUCT_TITLE_NOT_CONTAINS) {
    whereArray.push(
      " and segments.product_title NOT LIKE '%" +
      SETTINGS.PRODUCT_TITLE_NOT_CONTAINS[i].trim() +
      "%'"
    );
  }

  where += whereArray.join(" ");

  return where + " " + getFilterWhereString(SETTINGS);
}

// log(addCustomMetrics(row))
function addCustomMetricsToRow(row) {
  //CUSTOM_METRICS
  for (let metricName in CUSTOM_METRICS) {
    if (valueUndefined(row[CUSTOM_METRICS[metricName][1]]))
      throw "Error: Can't find the metric " + CUSTOM_METRICS[metricName][1];
    if (valueUndefined(row[CUSTOM_METRICS[metricName][0]]))
      throw "Error: Can't find the metric " + CUSTOM_METRICS[metricName][0];

    if (CUSTOM_METRICS[metricName][2] == "divide") {
      row[metricName] =
        row[CUSTOM_METRICS[metricName][0]] == 0 ||
          row[CUSTOM_METRICS[metricName][1]] == 0
          ? 0
          : round(
            row[CUSTOM_METRICS[metricName][0]] /
            row[CUSTOM_METRICS[metricName][1]],
            4
          );
    }
    if (CUSTOM_METRICS[metricName][2] == "multiply") {
      row[metricName] =
        row[CUSTOM_METRICS[metricName][0]] == 0 ||
          row[CUSTOM_METRICS[metricName][1]] == 0
          ? 0
          : round(
            row[CUSTOM_METRICS[metricName][0]] *
            row[CUSTOM_METRICS[metricName][1]],
            4
          );
    }
  }
  return row;
}

function getFilterMap(SETTINGS, row) {
  let filters = Object.keys(SETTINGS).map(function (x) {
    if (x.indexOf("FILTER") > -1) {
      return x;
    }
  });

  let filterMap = {};
  let filterParts = ["metric", "operator", "value"];
  let numberOfFilters = filters.length / filterParts.length;
  // log("number of filters: " + numberOfFilters)
  for (let i = 0; i < numberOfFilters; i++) {
    let filterName = "FILTER_" + (i + 1);
    if (Object.keys(SETTINGS).indexOf(filterName + "_METRIC") == -1) continue;
    if (SETTINGS[filterName + "_METRIC"] == "") continue;
    filterMap[filterName] = filterMap[filterName] || {};
    for (let x in filterParts) {
      filterMap[filterName][filterParts[x]] =
        SETTINGS[filterName + "_" + filterParts[x].toUpperCase()];
      if (
        filterParts[x] === "value" &&
        String(filterMap[filterName][filterParts[x]])
          .toLowerCase()
          .indexOf("avg") > -1
      ) {
        filterMap[filterName][filterParts[x]] = row.AverageCpc;
      }
    }
  }

  return filterMap;
}

//Whether to skip the entity based on the stats and filters
//Only check custom metrics e.g. cpa
function skipEntity(row, SETTINGS) {
  //custom metrics map
  //name : formula, whether high or low is good (e.g. ROAS high is good, CPA high is bad - used for Vs target calcs)
  let filterMap = getFilterMap(SETTINGS, row);

  // log(JSON.stringify(filterMap))

  for (let filter in filterMap) {
    let this_filter = filterMap[filter];
    if (filterNotInCustomMetrics(this_filter.metric)) continue;
    let eval_string =
      row[this_filter.metric] +
      " " +
      this_filter.operator +
      " " +
      this_filter.value;
    // log("eval_string: " + eval_string)
    if (!eval(eval_string)) {
      // log(JSON.stringify(row))
      // log("Skipping")
      // log(eval_string)
      return true;
    }
  }

  function filterNotInCustomMetrics(metric) {
    if (metric == 'ad_group_criterion.cpc_bid_micros') return false;
    return Object.keys(CUSTOM_METRICS).indexOf(metric) === -1;
  }

  return false;
}

function valueUndefined(val) {
  let result = false;

  if (val === 0) return false;
  if (!val) return true;
  if (typeof val === "undefined") return true;

  return result;
}

/**
 * Clear the output sheet
 * Only clear once everything has processed
 * @param {Object} SETTINGS 
 */
function clearOutputSheet(SETTINGS) {
  let sheetName = SETTINGS["TAB_NAME"];
  let sheet = SETTINGS.logSS.getSheetByName(sheetName);
  sheet
    .getRange(4, 1, sheet.getLastRow(), sheet.getLastColumn())
    .clearContent();
}

/*
 * Pull products from the api report
 * and add to the sheet for review
 */
function addProductsToSheet(SETTINGS) {
  let productChanges = getProductsFromApi(SETTINGS);

  let bidLogArray = [];
  // log(JSON.stringify(productChanges));
  for (let row_i in productChanges["rows"]) {
    let row = productChanges["rows"][row_i];

    let campaignName = SETTINGS.INCLUDE_CAMPAIGN ? row["campaign.name"] : "";
    let adGroupName = SETTINGS.INCLUDE_AD_GROUP ? row["ad_group.name"] : "";
    let campaignId = SETTINGS.INCLUDE_CAMPAIGN ? row["campaign.id"] : "";
    let adGroupId = SETTINGS.INCLUDE_AD_GROUP ? row["ad_group.id"] : "";

    let logRow = [
      SETTINGS["NAME"],
      row.AverageCpc,
      SETTINGS["N"],
      row['segments.product_item_id'],
      row['segments.product_title'],
      campaignName,
      adGroupName,
      row['metrics.clicks'],
      row['metrics.cost'],
      row['metrics.conversions'],
      row.Cpa,
      row.Roas,
      row.Cos,
      campaignId,
      adGroupId,
      SETTINGS.NOW,
    ];

    bidLogArray.push(logRow);
  }

  writeToSheet(SETTINGS, bidLogArray, SETTINGS["TAB_NAME"]);
  log(parseInt(bidLogArray.length) + " products added to the sheet");
  return bidLogArray.length > 0;
}

function shoppingPerformanceViewReport() {
  const cols = [
    'ad_group.id',
    'campaign.id',
    'campaign.name',
    'ad_group.name',
    'metrics.conversions_value',
    'metrics.impressions',
    'metrics.clicks',
    'metrics.cost_micros',
    'metrics.conversions',
    'ad_group_criterion.cpc_bid_micros',
    'campaign.bidding_strategy_type',
    'segments.product_item_id',
    'segments.product_title',
  ];
  let query = `SELECT ${cols.join(',')} FROM shopping_performance_view`
  let reportIter = AdsApp.report(query).rows();
  while (reportIter.hasNext()) {
    let row = reportIter.next();
    console.log(JSON.stringify(row))
  }
}

function assetGroupListingFilterReport() {
  const cols = [
    'ad_group.id',
    'campaign.id',
    'campaign.name',
    'ad_group.name',
    'metrics.conversions_value',
    'metrics.impressions',
    'metrics.clicks',
    'metrics.cost_micros',
    'metrics.conversions',
    'segments.product_item_id',
    'segments.product_title',
  ];
  let query = `SELECT ${cols.join(',')} FROM asset_group_listing_group_filter`
  let reportIter = AdsApp.report(query).rows();
  while (reportIter.hasNext()) {
    let row = reportIter.next();
    console.log(JSON.stringify(row))
  }
}


function getProductsFromApi(SETTINGS) {
  const cols = [
    'metrics.conversions_value',
    'metrics.impressions',
    'metrics.clicks',
    'metrics.cost_micros',
    'metrics.conversions',
    'segments.product_item_id',
    'segments.product_title',
  ];

  if (SETTINGS.INCLUDE_CAMPAIGN) {
    cols.push("campaign.status");
    cols.push("campaign.name");
    cols.push("campaign.id")
  }

  if (SETTINGS.INCLUDE_AD_GROUP) {
    cols.push("ad_group.status");
    cols.push("ad_group.name");
    cols.push("ad_group.id")
  }

  const reportName = "shopping_performance_view";

  let query = [
    "select",
    cols.join(","),
    "from",
    reportName,
    getQueryWhereString(SETTINGS),
    `and segments.date > ${SETTINGS.DATE_RANGE.split(",")[0]} and segments.date < ${SETTINGS.DATE_RANGE.split(",")[1]}`
  ].join(" ");

  log("product_group_view report query: " + query);

  let map = { ids: [], rows: {} };
  let report = AdsApp.report(query);
  exportReportToSheet(report, "Product Groups");
  let reportIter = report.rows();
  testLog(`totalNumEntities: ${reportIter.totalNumEntities()}`);
  if (!reportIter.hasNext()) {
    log("No products match the query");
  }

  let numberOfRows = 0;
  while (reportIter.hasNext()) {
    let row = reportIter.next();
    numberOfRows++;

    row['metrics.cost'] = row['metrics.cost_micros'] / 1000000
    row['metrics.impressions'] = parseInt(row['metrics.impressions'], 10);
    row['metrics.clicks'] = parseInt(row['metrics.clicks'], 10);
    row['metrics.conversions'] = parseFloat(
      row['metrics.conversions'].toString().replace(/,/g, "")
    );
    row['metrics.conversions_value'] = parseFloat(
      row['metrics.conversions_value'].toString().replace(/,/g, "")
    );

    row = addCustomMetricsToRow(row);
    // log(JSON.stringify(row))
    if (skipEntity(row, SETTINGS)) {
      // log(JSON.stringify(row))
      continue;
    }
    let row_id = String(numberOfRows) + row['product_item_id'];
    map["rows"][row_id] = {};
    map["rows"][row_id] = row;
  }

  //  log(JSON.stringify(map))
  log(numberOfRows + " initial products returned from the api query");
  return map;
}

function exportReportToSheet(report, sheetName) {
  const sheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(sheetName);
  if (!sheet) {
    return;
  }
  sheet.clear();
  if (!DEBUG) {
    sheet.getRange("A2").setValue("Report data will NOT be populated outside of debug mode");
    return;
  }
  report.exportToSheet(sheet);
}

function isToday(date) {
  return date.getDate() == new Date().getDate();
}

function writeArrayToSheet(array, sheet, start_row) {
  const rowsToInsert = array.length - (sheet.getLastRow() - 5)
  if (rowsToInsert > 0) {
    sheet.insertRowsAfter(5, rowsToInsert)
  }
  sheet.getRange(start_row, 1, array.length, array[0].length).setValues(array);
}

function getHeaderRow(SETTINGS) {

  return [
    'Rule Name',
    'Avg. CPC',
    'Look Back Period',
    'Item Id',
    'Title',
    'Campaign Name',
    'Ad Group Name',
    'Clicks',
    'Cost',
    'Conversions',
    'Cpa',
    'Roas',
    'COS%',
    'Campaign Id',
    'Adgroup ID',
    'Timestamp',
  ]

}

function writeToSheet(SETTINGS, logArray, sheetName) {
  if (!logArray.length) {
    return;
  }
  const spreadsheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
  new SheetFilter(spreadsheet, sheetName).removeFilter();

  const headerRow = getHeaderRow(SETTINGS)
  log(`header length: ${headerRow.length}`)
  log(`first log row length: ${logArray[0].length}`)
  logArray.unshift(headerRow);
  // log("Adding "+ logArray.length + " changes to " + sheetName)
  let sheet = SETTINGS.logSS.getSheetByName(sheetName);
  sheet
    .getRange("A2")
    .setValue(
      "Current Data For Lookback (" +
      SETTINGS.N +
      " days)" +
      " - " +
      SETTINGS.DATE_RANGE
    );
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  const startRow = 3
  sheet
    .getRange(startRow, 1, sheet.getLastRow(), sheet.getLastColumn())
    .clearContent();
  if (logArray.length === 1) {
    sheet.getRange("A2").setValue("No data found");
    return;
  }
  writeArrayToSheet(logArray, sheet, startRow);

  //sort by cost
  if (sheetName !== INPUT_TAB_NAME) {
    let costColumnNumber =
      parseInt(
        sheet
          .getRange(startRow, 1, 3, sheet.getLastColumn())
          .getValues()[0]
          .indexOf("Cost")
      ) + 1;
    sheet
      .getRange(startRow, 1, sheet.getLastRow(), sheet.getLastColumn())
      .sort({ column: costColumnNumber, ascending: false });
  }
  //   sheet.getRange(1,2,sheet.getLastRow(),sheet.getLastColumn()).setNumberFormat("0.00")
  //   sheet.getRange(1,13,sheet.getLastRow(),sheet.getLastColumn()).setNumberFormat("0.00%")
  //   sheet.getRange(1,7,sheet.getLastRow(),1).setNumberFormat("0.00%")
  //  sheet.getRange(1,10,sheet.getLastRow(),1).setNumberFormat("0.00%")

  new SheetFilter(spreadsheet, sheetName).createFilter();
}

function parseDateRange(SETTINGS) {
  let YESTERDAY = getAdWordsFormattedDate(1, "yyyyMMdd");

  SETTINGS.DATE_RANGE =
    getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd") + "," + YESTERDAY;
}

class SheetFilter {

  constructor(spreadsheet, sheetName) {
    this.spreadsheet = spreadsheet;
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
    const reportSheet = this.spreadsheet.getSheetByName(this.sheetName);
    if (reportSheet.getLastRow() < 2) {
      return;
    }
    const rangeValues = [3, 1, reportSheet.getLastRow(), reportSheet.getLastColumn()];
    const filterRange = reportSheet.getRange(...rangeValues);
    return filterRange;
  }
}

/*

    SETTINGS SECTION

    */

/**
 * Get the editors from the sheet
 * @param {drive element} - main control (settings) sheet
 * @param {int} - Number of the column containing the logs
 * @return {array} editors
 **/
function getEditorsFromSheet(CONTROL_SHEET, logsColumn) {
  let editors = CONTROL_SHEET.getRange(1, logsColumn).getValue();
  if (editors == "") {
    return;
  }
  if (editors.indexOf(",") > -1) {
    editors = editors.split(",");
    for (let e in editors) {
      editors[e] = editors[e].trim().toLowerCase();
    }
  } else {
    editors = [editors.trim().toLowerCase()];
  }
  return editors;
}

function getFilterHeaderTypes() {
  let map = {};
  for (let i = 1; i < NUMBER_OF_FILTERS + 1; i++) {
    map["FILTER_" + i + "_METRIC"] = "filter_metric";
    map["FILTER_" + i + "_OPERATOR"] = "filter_operator";
    map["FILTER_" + i + "_VALUE"] = "filter_value";
  }
  return map;
}

function getHeaderTypes() {
  let MCC_HEADER_TYPES = {};

  if (isMCC()) {
    MCC_HEADER_TYPES = { ID: "normal" };
  }

  let SINGLE_ACCOUNT_HEADER_TYPES = {
    NAME: "normal",
    EMAILS: "csv",
    FLAG: "bool",
    SEND_EMAIL: "bool",
    EMAIL_PREFIX: "normal",
    TAB_NAME: "normal",
    N: "normal",
    INCLUDE_CAMPAIGN: "bool",
    INCLUDE_AD_GROUP: "bool",
    CAMPAIGN_NAME_CONTAINS: "csv",
    CAMPAIGN_NAME_NOT_CONTAINS: "csv",
    ITEM_ID_CONTAINS: "csv",
    ITEM_ID_NOT_CONTAINS: "csv",
    PRODUCT_TITLE_CONTAINS: "csv",
    PRODUCT_TITLE_NOT_CONTAINS: "csv",
  };

  let FILTER_HEADER_TYPES = getFilterHeaderTypes();
  log(JSON.stringify(FILTER_HEADER_TYPES))

  let SINGLE_ACCOUNT_HEADER_TYPES2 = {
    LOG_SHEET_URL: "normal",
    LOGS_COLUMN: "normal",
  };

  let HEADER_TYPES = objectMerge(
    MCC_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES,
    FILTER_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES2
  );

  return HEADER_TYPES;
}
/**
 * Get the settings data minus headers
 * @returns {Object} - an object containing the header types
 */
function getSettingsData() {
  const controlSheet =
    SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME);
  let data = controlSheet
    .getDataRange()
    .getValues();
  data.shift();
  data.shift();
  data.shift();
  return data;
}

/**
 * get the logs column
 * @returns {void}
 */
function getLogsColumn() {
  const controlSheet =
    SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME);
  let logsColumn = 0;
  let col = 5;
  while (controlSheet.getRange(3, col).getValue()) {
    logsColumn = controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (logsColumn > 0) {
      break;
    }
    col++;
  }
  return logsColumn;
}

function scanForAccounts() {
  let controlSheet = 0;
  log("getting settings...");
  log(
    "The settings sheet should contain " + NUMBER_OF_FILTERS + " filter sets"
  );
  let map = {};

  const data = getSettingsData();

  const HEADER_TYPES = getHeaderTypes();

  log('HEADER_TYPES: ' + JSON.stringify(HEADER_TYPES))

  let HEADER = Object.keys(HEADER_TYPES);

  const logsColumn = getLogsColumn();

  let flagPosition = HEADER.indexOf("FLAG");
  console.log(`flagPosition: ${flagPosition}`)
  for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
    //if "run script" is not set to "yes", continue.
    let sheetRowNumber = rowIndex + 4;
    let row = data[rowIndex];

    if (String(row[0]) == "" || row[flagPosition].toLowerCase() != "yes") {
      continue;
    }
    let id = String(rowIndex);
    map[id] = { ROW_NUM: sheetRowNumber };
    for (let j in HEADER) {
      if (HEADER[j] == "LOGS_COLUMN") {
        map[id][HEADER[j]] = logsColumn;
        continue;
      }
      map[id][HEADER[j]] = row[j];
    }
  }
  let previousOperator = "";
  for (let id in map) {
    for (let key in map[id]) {
      let isFilterValue = key.indexOf("FILTER") > -1 && key.indexOf("VALUE") > -1;
      let filter_metric = !isFilterValue ? "" : map[id][key.replace("VALUE", "METRIC")];
      map[id][key] = processSetting(
        key,
        map[id][key],
        HEADER_TYPES,
        controlSheet,
        previousOperator,
        filter_metric
      );
      previousOperator = key.indexOf("OPERATOR") > -1 ? map[id][key] : "";
    }
  }
  return map;
}

function objectMerge() {
  for (let i = 1; i < arguments.length; i++)
    for (let a in arguments[i]) arguments[0][a] = arguments[i][a];
  return arguments[0];
}

function isList(operator) {
  let list_operators = [
    "IN",
    "NOT_IN",
    "CONTAINS_ANY",
    "CONTAINS_NONE",
    "CONTAINS_ALL",
  ];
  return list_operators.indexOf(operator) > -1;
}

// log(formatFilterValue("Example campaign name, another example", "CONTAINS_ANY"))
function formatFilterValue(value, operator, filter_metric) {
  // log(`formatting filter value. value: ${value}, operator: ${operator}, filter_metric: ${filter_metric}, `)
  if (h.isNumber(value)) return value;
  if (filter_metric === "ad_group_criterion.labels") {
    console.log("Swapping label string for resource name")
    value = String(value)
      .split(",")
      .map(function (x) {
        return getLabelResourceNameFromString(x);
      })
      .join(",");
  }
  if (String(value).split(",").length === 1) {
    return isList(operator) ? "['" + value + "']" : "'" + value + "'";
  }
  let arr = String(value).split(",");
  value =
    arr.length === 1
      ? arr[0]
      : "[" +
      arr
        .map(function (x) {
          return "'" + x.trim() + "'";
        })
        .join(",") +
      "]";
  return value;
}

// log(processSetting("FILTER_1", "world", {"FILTER_1":"filter_value"},"controlSheet"))
function processSetting(
  key,
  value,
  HEADER,
  controlSheet,
  previousOperator,
  filter_metric
) {
  if (key == "ROW_NUM") {
    return value;
  }
  if (String(value) === "") {
    return
  }
  let type = HEADER[key];

  switch (type) {
    case "filter_operator":
      // log(`filter operator: ${value}`)
      return value.replace("'", "");
    case "filter_value":
      value = formatFilterValue(value, previousOperator, filter_metric);
      return value;
    case "filter_metric":
      if (value === 'AdGroupName') {
        return 'ad_group.name'
      }
      if (value === 'Labels') {
        return 'ad_group_criterion.labels'
      }
      return convertAwqlMetricToGaql(value);
    case "label":
      return [
        controlSheet
          .getRange(3, Object.keys(HEADER).indexOf(key) + 1)
          .getValue(),
        value,
      ];
    case "normal":
      return value;
    case "bool":
      return value == "Yes" ? true : false;
    case "csv":
      let ret = value.split(",");
      ret = ret[0] == "" && ret.length == 1 ? [] : ret;
      if (ret.length == 0) {
        return [];
      } else {
        for (let r in ret) {
          ret[r] = String(ret[r]).trim();
        }
      }
      return ret;
    default:
      throw "error setting type " + type + " not recognised for " + key;
  }
}

/**
 * Take a report metric
 * Convert it from AWQl to GAQL
 * e.g. Clicks would become metrics.clicks
 * @param {String} field metric
 * @returns a GAQLmetric
 */
function convertAwqlMetricToGaql(field) {

  function pascalToSnake(str) {
    return String(str).replace(/[A-Z]/g, function (match, index) {
      return (index !== 0 ? '_' : '') + match.toLowerCase();
    });
  }

  const ignoreList = [
    'Ctr',
    'Roas',
    'Cos',
    'Cpa',
    'AverageCpc',
    'ConversionRate',
    'Rpc',
    'AdGroupName',
  ]
  if (ignoreList.indexOf(field) > -1) {
    return field
  }
  field = pascalToSnake(field)
  if (field === 'conversion_value') {
    field = 'conversions_value'
  }
  field = `metrics.${field}`
  return field
}

function processRowSettings(SETTINGS) {
  SETTINGS.NOW = Utilities.formatDate(
    new Date(),
    AdsApp.currentAccount().getTimeZone(),
    "MMM dd, yyyy HH:mm:ss"
  );

  SETTINGS.LOGS_COLUMN = getLogsColumn(SETTINGS.CONTROL_SHEET);
  SETTINGS.LOGS = [];
  SETTINGS.EMAILS_SHEET = true;

  let defaultNote =
    "Possible problems include: 1) There was an error (check the logs within Google Ads) 2) The script was stopped before completion";
  SETTINGS.CONTROL_SHEET.getRange(SETTINGS.ROW_NUM, SETTINGS.LOGS_COLUMN, 1, 1)
    .setValue(
      "The script is either still running or didn't finish successfully"
    )
    .setNote(defaultNote);

  parseDateRange(SETTINGS);
  SETTINGS.PREVIEW_MODE = AdsApp.getExecutionInfo().isPreview();

  createSheetIfNotExists(SETTINGS['TAB_NAME']);

}


function createSheetIfNotExists(sheetName) {
  const spreadsheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName)
    sheet.setName(sheetName);
  }
}

/**
 * Checks the settings for issues
 * @returns nothing
 **/
function checkSettings(SETTINGS) {
  //check the settings here
  if (SETTINGS.N === "") {
    log(JSON.stringify(SETTINGS));
    updateControlSheet(
      "Please set a lookback window. Note: Ensure the NUMBER_OF_FILTERS number is correct",
      SETTINGS
    );
  }

}


/**
 * Get AdWords Formatted date for n days back
 * @param {drive element} - drive element (such as folder or spreadsheet)
 * @param {drive element} - main control (settings) sheet
 * @param {int} - Number of the column containing the logs
 * @return nothing
 **/
function addEditors(spreadsheet, controlSheet, LOGS_COLUMN) {
  //check current editors, add if they don't exist
  let currentEditors = spreadsheet.getEditors();
  let currentEditorEmails = [];
  for (let c in currentEditors) {
    currentEditorEmails.push(currentEditors[c].getEmail().trim().toLowerCase());
  }

  let editors = controlSheet.getRange(1, LOGS_COLUMN).getValue();
  if (editors == "") {
    return;
  }
  if (editors.indexOf(",") > -1) {
    editors = editors.split(",");
    for (let e in editors) {
      editors[e] = editors[e].trim().toLowerCase();
    }
  } else {
    editors = [editors.trim().toLowerCase()];
  }

  for (let e in editors) {
    let index = currentEditorEmails.indexOf(editors[e]);
    if (currentEditorEmails.indexOf(editors[e]) == -1) {
      spreadsheet.addEditor(editors[e]);
    }
  }
}

function callBack() {
  // Do something here
  Logger.log("Finished");
}

function stringifyLogs(logs) {
  let s = "";
  for (let l in logs) {
    s += parseInt(l) + 1 + ") ";
    s += logs[l] + " ";
  }
  return s;
}

/**
 * Get AdWords Formatted date for n days back
 * @param {int} d - Numer of days to go back for start/end date
  * @return {String} - Formatted date yyyyMMdd
 **/
function getAdWordsFormattedDate(d, format) {
  let date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(
    date,
    AdsApp.currentAccount().getTimeZone(),
    format
  );
}

function log(message) {
  try {
    Logger.log(AdsApp.currentAccount().getName() + " - " + message);
  } catch (e) {
    console.log(message);
  }
}

function round(num, n) {
  return +(Math.round(parseFloat(num) + "e+" + n) + "e-" + n);
}

function main() {
  if (isMCC()) {
    let SETTINGS = scanForAccounts();
    //   log(JSON.stringify(SETTINGS))
    let ids = Object.keys(SETTINGS);
    if (ids.length == 0) {
      Logger.log("No Rules Specified");
      return;
    }
    MccApp.accounts()
      .withIds(ids)
      .withLimit(50)
      .executeInParallel("runRows", "callBack", JSON.stringify(SETTINGS));
  } else {
    let settings = scanForAccounts();

    //run all rows and all accounts
    for (let rowId in settings) {
      let rowSettings = settings[rowId]
      try {
        runScript(rowSettings);
      } catch (e) {
        console.error(e.stack)
      }
    }
  }
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

function runRows(INPUT) {
  log("running rows");
  let SETTINGS =
    JSON.parse(INPUT)[AdsApp.currentAccount().getCustomerId().toString()];
  for (let rowId in SETTINGS) {
    runScript(SETTINGS[rowId]);
  }
}

function Helper() {
  /**
   * Check if a string is a number, used when grabbing numbers from the sheet
   * If the string contains anything but numbers and a full stop (.) it returns false
   * @param {number as a string}
   * @returns {bool}
   **/
  this.isNumber = function (n) {
    if (typeof n == "number") return true;
    n = n.trim();
    let digits = n.split("");
    for (let d in digits) {
      if (digits[d] == ".") {
        continue;
      }
      if (isNaN(digits[d])) {
        return false;
      }
    }
    return true;
  };

  /**
   * Calculate ROAS
   * @param {number} - Conv. Value
   * @param {number} - metrics.cost
   * @returns {number}
   **/
  this.calculateRoas = function (conversions_value, cost) {
    if (cost == 0) return 0;
    if (conversions_value == 0) return 0;
    if (cost > conversions_value) return 0;
    return conversions_value / cost;
  };

  /**
   * Return the column number of the logs column
   * @param {google sheet} control/settings sheet
   * @return {number} - Logs column
   **/
  this.getLogsColumn = function (controlSheet) {
    let col = 5;
    let LOGS_COLUMN = 0;
    while (String(controlSheet.getRange(3, col).getValue())) {
      LOGS_COLUMN =
        controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
      if (LOGS_COLUMN > 0) {
        break;
      }
      col++;
    }
    return LOGS_COLUMN;
  };

  /**
   * Turn an array of logs into a numbered string
   * @param {array} logs
   * @return {String} - Logs
   **/
  this.stringifyLogs = function (logs) {
    let s = "";
    for (let l in logs) {
      s += parseInt(l) + 1 + ") ";
      s += logs[l] + " ";
    }
    return s;
  };

  /**
   * Get AdWords Formatted date for n days back
   * @param {int} d - Numer of days to go back for start/end date
   * @return {String} - Formatted date yyyyMMdd
   **/
  this.getAdWordsFormattedDate = function (d, format) {
    let date = new Date();
    date.setDate(date.getDate() - d);
    return Utilities.formatDate(
      date,
      AdsApp.currentAccount().getTimeZone(),
      format
    );
  };

  this.round = function (num, n) {
    return +(Math.round(num + "e+" + n) + "e-" + n);
  };

  /**
   * Add editors to the sheet
   * @param {drive element} - drive element (such as folder or spreadsheet)
   * @param {array} - editors to add
   * @return nothing
   **/
  this.addEditors = function (spreadsheet, editors) {
    //check current editors, add if they don't exist
    let currentEditors = spreadsheet.getEditors();
    let currentEditorEmails = [];
    for (let c in currentEditors) {
      currentEditorEmails.push(
        currentEditors[c].getEmail().trim().toLowerCase()
      );
    }

    for (let e in editors) {
      if (currentEditorEmails.indexOf(editors[e]) == -1) {
        spreadsheet.addEditor(editors[e]);
      }
    }
  };
}

//uses helpers.js

function Setting() {
  this.parseDateRange = function (SETTINGS) {
    let YESTERDAY = h.getAdWordsFormattedDate(1, "yyyyMMdd");
    SETTINGS.DATE_RANGE = "20000101," + YESTERDAY;

    if (SETTINGS.DATE_RANGE_LITERAL == "LAST_N_DAYS") {
      SETTINGS.DATE_RANGE =
        h.getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd") + "," + YESTERDAY;
    }

    if (SETTINGS.DATE_RANGE_LITERAL == "LAST_N_MONTHS") {
      let now = new Date(
        Utilities.formatDate(
          new Date(),
          AdsApp.currentAccount().getTimeZone(),
          "MMM dd, yyyy HH:mm:ss"
        )
      );
      now.setHours(12);
      now.setDate(0);

      let TO = Utilities.formatDate(now, "PST", "yyyyMMdd");
      now.setDate(1);

      let counter = 1;
      while (counter < SETTINGS.N) {
        now.setMonth(now.getMonth() - 1);
        counter++;
      }

      let FROM = Utilities.formatDate(now, "PST", "yyyyMMdd");
      SETTINGS.DATE_RANGE = FROM + "," + TO;
    }
  };
}

function testLog() {
  if (!DEBUG) {
    return;
  }
  console.log(...arguments);
}