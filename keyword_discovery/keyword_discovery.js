/**
 * Keyword Discovery Script
 * @author Charles Bannister
 * Surface new keyword opportunities
 * Purchased from shabba.io and not for resale or redistribution - thanks!
 * @version 1.0.0
 * Free updates & support at https://shabba.io/script/9
 * 
**/

// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1yCg93xb3Yhx1c8q9AIbA9KHfa5cyMIquzvA5GW3OOH4/copy
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";

try {
  module.exports = {
    containsStringsInSearchTerm,
    containsStringsNotInSearchTerm,
    fuzzyMatchScore
  };
} catch (e) {
  console.log("")
}



const INPUT_TAB_NAME = "Settings";

const SCRIPT_NAME = "Keyword Discovery"
const IGNORED_TERMS_SHEET_NAME = "Ignored Terms"

const NUMBER_OF_FILTERS = 6;

//set to true to send emails regardless of preview status
const OVERRIDE_PREVIEW_EMAIL = false;

//No need to edit anything below this line

let h = new Helper();
let s = new Setting();

const OPTIONS = { includeZeroImpressions: true };

const SINGLE_ACCOUNT_HEADER_TYPES = {
  NAME: "normal",
  EMAILS: "csv",
  FLAG: "bool",
  EMAIL_ALERT: "bool",
  TAB_NAME: "normal",
  N: "normal",
  SEARCH_TERM_CONTAINS: 'csv',
  SEARCH_TERM_NOT_CONTAINS: 'csv',
  FUZZY_THRESHOLD: "normal",
};

const SINGLE_ACCOUNT_HEADER_TYPES_AFTER = {
  LOG_SHEET_URL: "normal",
  LOGS_COLUMN: "normal",
};

const CUSTOM_METRICS = {
  //metric, metric, operator (divide or multiply, whether high or low is good)
  Ctr: ["Clicks", "Impressions", "divide", "high"],
  Roas: ["ConversionValue", "Cost", "divide", "high"],
  Cos: ["Cost", "ConversionValue", "divide", "low"],
  Cpa: ["Cost", "Conversions", "divide", "low"],
  AverageCpc: ["Cost", "Clicks", "divide", "low"],
  ConversionRate: ["Conversions", "Clicks", "divide", "high"],
  Rpc: ["ConversionValue", "Clicks", "divide", "high"],
};

function runScript(SETTINGS) {
  log("Script Started");

  SETTINGS.SPREADSHEET = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
  SETTINGS.CONTROL_SHEET = SETTINGS.SPREADSHEET.getSheetByName(INPUT_TAB_NAME);

  checkSettings(SETTINGS);

  processRowSettings(SETTINGS);

  log(`This row's settings: ${JSON.stringify(SETTINGS)}`);

  let searchTermsToIgnore = getSearchTermsToIgnoreFromSheet(SETTINGS);

  // log(`searchTermsToIgnore: ${searchTermsToIgnore}`);

  addIgnoredSearchTermsToSheet(searchTermsToIgnore, SETTINGS);
  let keywords = getKeywords();
  // log(`keywords: ${keywords}`);
  populateOutputSheet(keywords, SETTINGS);

  SETTINGS.LOGS.push("The script ran successfully");
  updateControlSheet("", SETTINGS);

  log("Finished row " + SETTINGS.ROW_NUM);
}

/**
 * query the api for search terms
 * populate the sheet
 * don't add search terms from the ignored list
 * @param {Object} SETTINGS
 */
function populateOutputSheet(keywords, SETTINGS) {
  log("clearing output sheet");
  let sheet_operations = new SheetOperations(SETTINGS);
  let sheet = sheet_operations.getSheetByName(SETTINGS["TAB_NAME"]);
  if (sheet.getMaxRows() > 3) {
    sheet.deleteRows(4, parseInt(sheet.getMaxRows()) - 3);
  }

  SETTINGS["FILTER_MAP"] = getFilterMap(SETTINGS);

  let cols = [
    "AdGroupId",
    "CampaignId",
    "CampaignName",
    "AdGroupName",
    "Query",
    "ConversionValue",
    "Impressions",
    "Clicks",
    "Cost",
    "Conversions",
    "TopImpressionPercentage",
    "AbsoluteTopImpressionPercentage",
    "QueryTargetingStatus",
  ];

  let reportName = "SEARCH_QUERY_PERFORMANCE_REPORT";

  let query = [
    "select",
    cols.join(","),
    "from",
    reportName,
    getQueryWhereString(SETTINGS),
    "during",
    SETTINGS.DATE_RANGE,
  ].join(" ");

  log(reportName + " query: " + query);

  let logArray = [];

  let reportIter = AdWordsApp.report(query, OPTIONS).rows();
  if (!reportIter.hasNext()) {
    log("No search terms found (initial API query)")
  }
  // AdWordsApp.report(query, OPTIONS).exportToSheet(SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName("Sheet5"))
  while (reportIter.hasNext()) {
    let row = reportIter.next();

    row.Impressions = parseInt(row.Impressions, 10);
    row.Clicks = parseInt(row.Clicks, 10);
    row.Conversions = parseFloat(row.Conversions.toString().replace(/,/g, ""));

    row.Cost = parseFloat(row.Cost.toString().replace(/,/g, ""));
    row.ConversionValue = parseFloat(
      row.ConversionValue.toString().replace(/,/g, "")
    );

    row = addCustomMetricsToRow(row);

    if (skipEntity(row, SETTINGS)) {
      continue;
    }
    if (!containsStringsInSearchTerm(row.Query, SETTINGS)) {
      // console.log(`skipping due to contains string logic`)
      continue;
    }
    if (!containsStringsNotInSearchTerm(row.Query, SETTINGS)) {
      // console.log(`skipping due to NOT contains string logic`)
      continue;
    }
    if (fuzzyStringInArray(row.Query, keywords, SETTINGS['FUZZY_THRESHOLD'])) {
      continue;
    }


    let logRow = [
      row.Query,
      row.QueryTargetingStatus,
      false,
      row.Impressions,
      row.Clicks,
      row.Ctr,
      row.Cost,
      row.Conversions,
      row.Cpa,
      row.ConversionRate,
      row.ConversionValue,
      row.Roas,
    ];

    logArray.push(logRow);
  }

  log(logArray.length + " queries found");
  if (logArray.length === 0) {
    return
  };

  sheet_operations.writeToSheet(logArray, SETTINGS["TAB_NAME"], 4);

  postWriteOperations(logArray, SETTINGS);

  if (SETTINGS["EMAIL_ALERT"]) {
    sendEmailAlert(SETTINGS, logArray.length);
  }
}

/**
 * check if the search term contains the required strings
 * @param {string} searchTerm
 * @param {object} SETTINGS
 */
function containsStringsInSearchTerm(searchTermToCheck, SETTINGS) {
  if (SETTINGS.SEARCH_TERM_CONTAINS.length === 0) {
    return true;
  }
  for (let searchTermContainsString of SETTINGS.SEARCH_TERM_CONTAINS) {
    if (searchTermToCheck.indexOf(searchTermContainsString) === -1) {
      return false;
    }
  }
  return true;
}

/**
 * check the search term does not contain the required strings
 * @param {string} searchTerm
 * @param {object} SETTINGS
 */
function containsStringsNotInSearchTerm(searchTermToCheck, SETTINGS) {
  if (SETTINGS.SEARCH_TERM_NOT_CONTAINS.length === 0) {
    return true;
  }
  for (let searchTermNotContainsString of SETTINGS.SEARCH_TERM_NOT_CONTAINS) {
    if (searchTermToCheck.indexOf(searchTermNotContainsString) > -1) {
      return false;
    }
  }
  return true;
}

function getKeywords() {
  const MIN_KEYWORD_IMPRESSIONS = 0;
  const DATE_RANGE = 'LAST_30_DAYS'
  let query = "SELECT CampaignId, AdGroupId, Criteria, KeywordMatchType " +
    "FROM   KEYWORDS_PERFORMANCE_REPORT " +
    "WHERE Status = ENABLED AND IsNegative = FALSE AND Impressions > " +
    MIN_KEYWORD_IMPRESSIONS +
    "DURING " +
    DATE_RANGE
  var keywordReport = AdsApp.report(query);

  var keywordRows = keywordReport.rows();
  let keywords = [];
  while (keywordRows.hasNext()) {
    var row = keywordRows.next();
    keywords.push(row.Criteria);
  }
  return keywords;
}


/**
 *
 * @param {object} SETTINGS
 * @returns {array} queries to ignore
 */
function getSearchTermsToIgnoreFromSheet(SETTINGS) {

  let outputSheet = SETTINGS.OUTPUT_SHEET;

  let data = outputSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  let header = data.shift();
  // log(`header: ${header}`)
  let searchTermsToIgnore = []

  for (let dataIndex in data) {
    let row = data[dataIndex];

    let query = String(row[header.indexOf("Search Term")]);

    let this_is_ok = row[header.indexOf("Ignore")];
    // log(`this_is_ok: ${this_is_ok}, query: ${query}`)
    if (!this_is_ok) {
      continue
    }

    searchTermsToIgnore.push(query)

  }

  return searchTermsToIgnore;
}

function sendEmailAlert(SETTINGS, numberOfSearchTerms) {
  if (SETTINGS.PREVIEW_MODE && !OVERRIDE_PREVIEW_EMAIL) return;

  SETTINGS.LOGS.push(`Rule name: ${SETTINGS["NAME"]}`);
  SETTINGS.LOGS.push(`${numberOfSearchTerms} search terms matched the rule`);

  //Send email
  let SUB = AdWordsApp.currentAccount().getName() + " - New Keyword Discovery Alert";
  let MSG =
    `Hi,<br><br>The ${SCRIPT_NAME} script has results. Here are the logs:<br><br>`;

  MSG += "<ul>";
  for (let l in SETTINGS.LOGS) {
    MSG += "<li>";
    MSG += SETTINGS.LOGS[l];
    MSG += "</li>";
  }
  MSG += "</ul>";

  MSG +=
    "<br><br>Visit the Google Sheet to review the terms:<br>" +
    SETTINGS.LOG_SHEET_URL;
  MSG += "<br><br>Thanks,";
  MSG += "<br><br>Charles";
  MSG += "<br><br><br><a href='https://shabba.io'>shabba.io</a>";
  let emails = SETTINGS.EMAILS;

  for (let i in emails) {
    MailApp.sendEmail({
      to: emails[i],
      subject: SUB,
      htmlBody: MSG,
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
    throw errorMessage;
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
    // log(key + " - " + SETTINGS[key]);
    let filter_value = SETTINGS[key.replace("METRIC", "VALUE")];
    let value_split = String(filter_value).split(",");
    for (let v in value_split) {
      value_split[v] = getLabelIdFromString(value_split[v]);
    }
    SETTINGS[key.replace("METRIC", "VALUE")] = value_split.join(",");
  }
  return SETTINGS;
}

function getFilterWhereString(SETTINGS) {
  //custom metrics map
  //name : formula, whether high or low is good (e.g. ROAS high is good, CPA high is bad - used for Vs target calcs)

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
    }
  }

  let where = "";

  let whereArray = [];

  //turn the filters object into a where statement string
  whereArray.push(filtersToWhereStatement(filterMap, filterParts));
  function filtersToWhereStatement(filterMap, filterParts) {
    let str = [];
    for (let filter in filterMap) {
      //if the metric is a custom metric, continue
      if (Object.keys(CUSTOM_METRICS).indexOf(filterMap[filter]["metric"]) > -1)
        continue;

      str.push("and");
      for (let p in filterParts) {
        str.push(filterMap[filter][filterParts[p]]);
      }
    }
    return str.join(" ");
  }

  where += whereArray.join(" ");

  return where;
}

function getLabelIdFromString(str) {
  let labels = AdsApp.labels()
    .withCondition("Name = '" + str.trim() + "'")
    .get();
  if (labels.hasNext()) {
    let label = labels.next();
    //  log(label.getId())
    return label.getId();
  } else {
    checkLabel(str);
    return "0000";
  }
}

//if the label doesn't exist, create it
function checkLabel(labelName) {
  if (labelName == "") {
    return;
  }

  let labels = AdWordsApp.labels().get();
  let exists = false;
  while (labels.hasNext()) {
    let label = labels.next();
    if (label.getName() == labelName) {
      exists = true;
    }
  }
  if (!exists) {
    AdWordsApp.createLabel(labelName);
  }
}

function getQueryWhereString(SETTINGS) {
  let where = "where CampaignStatus = ENABLED and AdGroupStatus = ENABLED ";
  if (SETTINGS.QUERY_TARGETING_STATUS) {
    where += " and QueryTargetingStatus = " + SETTINGS.QUERY_TARGETING_STATUS;
  }

  let whereArray = [];

  for (let i in SETTINGS.CAMPAIGN_NAME_CONTAINS) {
    whereArray.push(
      " and CampaignName CONTAINS_IGNORE_CASE '" +
      SETTINGS.CAMPAIGN_NAME_CONTAINS[i].trim() +
      "'"
    );
  }

  for (let i in SETTINGS.CAMPAIGN_NAME_NOT_CONTAINS) {
    whereArray.push(
      " and CampaignName DOES_NOT_CONTAIN_IGNORE_CASE '" +
      SETTINGS.CAMPAIGN_NAME_NOT_CONTAINS[i].trim() +
      "'"
    );
  }

  where += whereArray.join(" ");

  return where + " " + getFilterWhereString(SETTINGS);
}

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
          : roundFloatToTwoDecimalPlaces(
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
          : roundFloatToTwoDecimalPlaces(
            row[CUSTOM_METRICS[metricName][0]] *
            row[CUSTOM_METRICS[metricName][1]],
            4
          );
    }
  }
  return row;
}

function postWriteOperations(logArray, SETTINGS) {
  let sheet_operations = new SheetOperations(SETTINGS);
  let sheet = sheet_operations.getSheetByName(SETTINGS["TAB_NAME"]);
  sheet
    .getRange("A2")
    .setValue(
      "Current Data For Lookback (" +
      SETTINGS.N +
      " days)" +
      " - " +
      SETTINGS.DATE_RANGE
    );
  let header = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues();
  addValidation(logArray.length, sheet, header, sheet_operations);
  sheet_operations.sortColumn(sheet, parseInt(header[0].indexOf("Cost")) + 1);
}

function addValidation(rows, sheet, header, sheet_operations) {
  //add validaiton to the keywords sheet
  let checkbox_cols = [];
  for (let value_index in header[0]) {
    let value = header[0][value_index];
    if (value == "Ignore") {
      checkbox_cols.push(parseInt(value_index) + 1)
    };
  }
  let start_row = 4;
  // log(`checkbox_cols: ${checkbox_cols}`)
  for (let checkbox_cols_index in checkbox_cols) {
    sheet_operations.addCheckBox(
      start_row,
      checkbox_cols[checkbox_cols_index],
      rows,
      sheet
    );
  }

}

class SheetOperations {
  constructor(SETTINGS) {
    this.SETTINGS = SETTINGS;
  }

  addDropdownValidation(validation_list, sheet, start_row, rows, column) {
    let rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(validation_list)
      .build();
    sheet.getRange(start_row, column, rows, 1).setDataValidation(rule);
  }

  addCheckBox(start_row, column, rows, sheet) {
    let range = sheet.getRange(start_row, column, rows, 1);

    let enforceCheckbox = SpreadsheetApp.newDataValidation();
    enforceCheckbox.requireCheckbox();
    enforceCheckbox.setAllowInvalid(false);
    enforceCheckbox.build();

    range.setDataValidation(enforceCheckbox);
  }

  getSheetByName(name) {
    // console.log("Getting sheet by name: " + name)
    try {
      return SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(name);
    } catch (e) {
      let sheets = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheets();
      let sheet_names = [];
      for (let sheet_key in sheets) {
        let sheet = sheets[sheet_key];
        sheet_names.push(sheet.getName());
      }

      log("There was a problem getting sheet '" + name + "'");
      log("Here are the availble sheets: " + sheet_names.join(", "));
      log("And here's the error: ");
      throw e;
    }
  }

  appendToSheet(logArray, sheetName, start_row) {
    if (logArray.length === 0) return;

    let sheet = this.getSheetByName(sheetName);

    sheet.insertRowsAfter(start_row - 1, logArray.length);

    let max_rows = 20000;

    if (sheet.getLastRow() > max_rows) {
      sheet.deleteRows(max_rows, sheet.getLastRow() - max_rows);
    }

    sheet
      .getRange(start_row, 1, logArray.length, logArray[0].length)
      .setValues(logArray);
  }

  getAllRowsData(sheet, start_row, num_columns) {
    let start_column = 1;
    return sheet
      .getRange(start_row, start_column, sheet.getLastRow(), num_columns)
      .getValues();
  }

  writeToSheet(logArray, sheetName, start_row) {
    log("Writing to " + sheetName);
    if (logArray.length === 0) return;

    let sheet = this.getSheetByName(sheetName);
    let lastRow = this.getLastRow(sheet, 4);
    // log("lastRow: " +  lastRow)
    if (lastRow < parseInt(logArray.length) + 20) {
      sheet.insertRowsAfter(lastRow, lastRow);
    }

    log("writing to sheet");
    sheet
      .getRange(4, 1, logArray.length, logArray[0].length)
      .setValues(logArray);
    // Utilities.sleep(50000)
  }

  clearRange(sheet, row, column, rows, columns) {
    let values = [];
    for (let r = 0; r < rows; r++) {
      let this_row = [];
      for (let c = 0; c < columns; c++) {
        this_row.push("");
      }
      values.push(this_row);
    }
    sheet.getRange(row, column, rows, columns).setValues(values);
  }

  /*Get the last row based on the first column's values */
  getLastRow(sheet, startRow) {
    let row = startRow;
    while (sheet.getRange(row, 1).getValue()) {
      row++;
    }
    row = row - 1;
    if (row < 1) return 1;
    return row;
  }

  addNumberFormat(sheet_name, column, start_row) {
    //   sheet.getRange(1,2,sheet.getLastRow(),sheet.getLastColumn()).setNumberFormat("0.00")
    //   sheet.getRange(1,13,sheet.getLastRow(),sheet.getLastColumn()).setNumberFormat("0.00%")
    //   sheet.getRange(1,7,sheet.getLastRow(),1).setNumberFormat("0.00%")
    //  sheet.getRange(1,10,sheet.getLastRow(),1).setNumberFormat("0.00%")
  }

  sortColumn(sheet, columnNumber) {
    let startRow = 4;
    if (sheet.getLastRow() - startRow < 1) {
      return
    }
    sheet
      .getRange(
        startRow,
        1,
        sheet.getLastRow() - startRow,
        sheet.getLastColumn()
      )
      .sort({ column: columnNumber, ascending: false });
  }
}

function getFilterMap(SETTINGS) {
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
    if (SETTINGS[filterName + "_METRIC"] == "") {
      continue;
    }
    filterMap[filterName] = filterMap[filterName] || {};
    for (let x in filterParts) {
      filterMap[filterName][filterParts[x]] =
        SETTINGS[filterName + "_" + filterParts[x].toUpperCase()];
    }
  }

  return filterMap;
}

//Whether to skip the entity based on the stats and filters
//Only check custom metrics e.g. cpa
function skipEntity(row, SETTINGS) {
  //custom metrics map
  //name : formula, whether high or low is good (e.g. ROAS high is good, CPA high is bad - used for Vs target calcs)
  let filterMap = SETTINGS["FILTER_MAP"];

  // log(JSON.stringify(filterMap))

  // log(JSON.stringify(row))

  for (let filter in filterMap) {
    let this_filter = filterMap[filter];
    if (filterNotInCustomMetrics(this_filter.metric)) continue;
    let eval_string =
      row[this_filter.metric] +
      " " +
      this_filter.operator +
      " " +
      String(this_filter.value);
    try {
      if (!eval(eval_string)) {
        return true;
      }
    } catch (e) {
      log("eval_string: " + eval_string);
      throw e;
    }
  }

  function filterNotInCustomMetrics(metric) {
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

function calculateCpaBid(bid, target, actual, min) {
  if (actual === 0) return min;
  return bid + bid * getChangePercentage(actual, target);
}

function calculateRoasBid(bid, target, actual, min) {
  if (actual === 0) return min;
  return bid + bid * getChangePercentage(target, actual);
}

/**
 * add positive and negative keywords
 * @returns nothing
 **/
function addIgnoredSearchTermsToSheet(searchTermsToIgnore, SETTINGS) {
  if (searchTermsToIgnore.length === 0) {
    return;
  }

  let logArray = [];
  //for each search term add to the logArray array along with a date
  for (let i in searchTermsToIgnore) {
    let row = [];
    row.push(searchTermsToIgnore[i]);
    row.push(SETTINGS.NOW);
    logArray.push(row);
  }

  const sheet_operations = new SheetOperations(SETTINGS);

  sheet_operations.appendToSheet(logArray, IGNORED_TERMS_SHEET_NAME, 5);
  SETTINGS.LOGS.push(
    parseInt(searchTermsToIgnore.length) + " search terms added to the ignore list"
  );

}

function addMatchType(text, matchType) {
  matchType = matchType.toLowerCase();
  if (matchType === "exact") return "[" + text + "]";
  if (matchType === "phrase") return '"' + text + '"';
  if (matchType === "broad") return text;
  throw "Error: Match type not recognised";
  //text.replace("[","").replace("]","").toLowerCase();
}

function isToday(date) {
  return date.getDate() == new Date().getDate();
}

//update the change log
//return false if the Id has already been updated today
function updateChangeLog(CampaignName, AdGroupName, Id, SETTINGS) {
  //grab the change log
  //if the id does not exist in the sheet, update the sheet and return true
  //if the id exists and it's today, return false
  //if the id exists but it's not today, update the sheet and return true

  if (AdsApp.getExecutionInfo().isPreview()) return true;

  let map = {};

  let sheet = SETTINGS.CHANGE_LOG_SHEET;
  if (sheet.getLastRow() < 2) {
    log("No change log values");
    map = updateMap(map, CampaignName, AdGroupName, Id, SETTINGS.NOW);
    writeMapToSheet(sheet, map);
    return true;
  }

  //grab the current details and update
  let data = sheet.getDataRange().getValues();
  data.shift();
  for (let d in data) {
    let row = data[d];
    let campaign = row[0];
    let adgroup = row[1];
    let id = row[2];
    let date = row[3];
    map[campaign + adgroup + id] = {};
    map[campaign + adgroup + id]["campaign"] = campaign;
    map[campaign + adgroup + id]["adgroup"] = adgroup;
    map[campaign + adgroup + id]["id"] = id;
    map[campaign + adgroup + id]["date"] = date;
  }

  let this_id = CampaignName + AdGroupName + Id;

  if (valueUndefined(map[this_id])) {
    // log("Change log id undefined: " + this_id)
    // log(JSON.stringify(map))
    map = updateMap(map, CampaignName, AdGroupName, Id, SETTINGS.NOW);
    writeMapToSheet(sheet, map);
    return true;
  }

  let this_date = new Date(map[this_id]["date"]);

  //don't update the bid if the log contains today's date
  if (isToday(this_date)) {
    // log("Id already updated today")
    return false; //do not update
  } else {
    // log("Value updated but it wasn't today")
    map = updateMap(map, CampaignName, AdGroupName, Id, SETTINGS.NOW);
    writeMapToSheet(sheet, map);
    return true;
  }

  function writeMapToSheet(sheet, map) {
    // log("Adding " + Object.keys(map).length + " rows to the change log")
    // log("Writing this map: " + JSON.stringify(map))
    sheet.clear();
    let header = ["Campaign", "Ad group", "Id", "Last Updated"];
    let logArray = [header];
    for (let id in map) {
      logArray.push([
        map[id]["campaign"],
        map[id]["adgroup"],
        map[id]["id"],
        map[id]["date"],
      ]);
    }

    writeArrayToSheet(logArray, sheet, 1);
  }

  function updateMap(map, CampaignName, AdGroupName, Id, NOW) {
    let this_id = CampaignName + AdGroupName + Id;
    map[this_id] = {};
    map[this_id]["campaign"] = CampaignName;
    map[this_id]["adgroup"] = AdGroupName;
    map[this_id]["id"] = Id;
    map[this_id]["date"] = NOW;
    return map;
  }
}

function getChangePercentage(before, after) {
  if (before === 0 && after === 0) return 0;
  return (after - before) / before;
}

function parseDateRange(SETTINGS) {
  let YESTERDAY = getAdWordsFormattedDate(1, "yyyyMMdd");

  SETTINGS.DATE_RANGE =
    getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd") + "," + YESTERDAY;
}

function getActualVsTarget(target_type, actual, target) {
  if (target_type === "CPA") {
    return target / actual;
  }

  return actual / target;
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
    map["FILTER_" + i + "_METRIC"] = "normal";
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

  let FILTER_HEADER_TYPES = getFilterHeaderTypes();

  let HEADER_TYPES = objectMerge(
    MCC_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES,
    FILTER_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES_AFTER
  );

  return HEADER_TYPES;
}

function buildTestSettings() {
  let map = {};
  for (let key in SINGLE_ACCOUNT_HEADER_TYPES) {
    map[key] = "random value";
  }
  return map;
}

function scanForAccounts() {
  log("getting settings...");
  log(
    "The settings sheet should contain " + NUMBER_OF_FILTERS + " filter sets"
  );
  let map = {};
  let controlSheet =
    SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME);
  let data = controlSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();

  HEADER_TYPES = getHeaderTypes();

  let HEADER = Object.keys(HEADER_TYPES);

  let LOGS_COLUMN = 0;
  let col = 5;
  while (controlSheet.getRange(3, col).getValue()) {
    LOGS_COLUMN = controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) {
      break;
    }
    col++;
  }

  // log(HEADER)
  let flagPosition = HEADER.indexOf("FLAG");

  for (let k in data) {
    //if "run script" is not set to "yes", continue.
    // log(data[k][flagPosition])
    if (data[k][0] == "" || data[k][flagPosition].toLowerCase() != "yes") {
      continue;
    }
    let rowNum = parseInt(k, 10) + 4;
    let id = data[k][0];
    let rowId = id + "/" + rowNum;
    map[id] = map[id] || {};
    map[id][rowId] = { ROW_NUM: parseInt(k, 10) + 4 };
    for (let j in HEADER) {
      if (HEADER[j] == "LOGS_COLUMN") {
        map[id][rowId][HEADER[j]] = LOGS_COLUMN;
        continue;
      }
      map[id][rowId][HEADER[j]] = data[k][j];
    }
  }

  if (Object.keys(map).length == 0) {
    console.warn("No rules were specified");
    console.warn(
      "To run the script, add settings and set 'Run Script' to 'Yes'"
    );
    return;
  }

  let previousOperator = "";
  for (let id in map) {
    for (let rowId in map[id]) {
      for (let key in map[id][rowId]) {
        let isFilterValue =
          key.indexOf("FILTER") > -1 && key.indexOf("VALUE") > -1;
        let filter_metric = !isFilterValue
          ? ""
          : map[id][rowId][key.replace("VALUE", "METRIC")];
        map[id][rowId][key] = processSetting(
          key,
          map[id][rowId][key],
          HEADER_TYPES,
          controlSheet,
          previousOperator,
          filter_metric
        );
        previousOperator =
          key.indexOf("OPERATOR") > -1 ? map[id][rowId][key] : "";
      }
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
  if (h.isNumber(value)) return value;
  if (filter_metric === "Labels") {
    value = String(value)
      .split(",")
      .map(function (x) {
        return getLabelIdFromString(x);
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
  let type = HEADER[key];
  if (key == "ROW_NUM") {
    return value;
  }
  switch (type) {
    case "filter_operator":
      return value.replace("'", "");
      break;
    case "filter_value":
      value = formatFilterValue(value, previousOperator, filter_metric);
      return value;
      break;
    case "label":
      return [
        controlSheet
          .getRange(3, Object.keys(HEADER).indexOf(key) + 1)
          .getValue(),
        value,
      ];
      break;
    case "normal":
      return value;
      break;
    case "bool":
      return value == "Yes" ? true : false;
      break;
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
      break;
    default:
      throw "error setting type " + type + " not recognised for " + key;
  }
}

function processRowSettings(SETTINGS) {
  SETTINGS.NOW = Utilities.formatDate(
    new Date(),
    AdWordsApp.currentAccount().getTimeZone(),
    "MMM dd, yyyy HH:mm:ss"
  );
  SETTINGS["QUERY_TARGETING_STATUS"] =
    SETTINGS["QUERY_TARGETING_STATUS"] == "ALL"
      ? ""
      : SETTINGS["QUERY_TARGETING_STATUS"];

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
  SETTINGS.PREVIEW_MODE = AdWordsApp.getExecutionInfo().isPreview();

  if (SETTINGS.PREVIEW_MODE) {
    let msg = "Running in preview mode. No changes will be made.";
    SETTINGS.LOGS.push(msg);
  }

  const spreadsheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
  SETTINGS.SPREADSHEET = spreadsheet
  SETTINGS.OUTPUT_SHEET = spreadsheet.getSheetByName(String(SETTINGS.TAB_NAME));
  SETTINGS.IGNORED_TERMS_SHEET = spreadsheet.getSheetByName(IGNORED_TERMS_SHEET_NAME);
}

/**
 * Checks the settings for issues
 * @returns nothing
 **/
function checkSettings(SETTINGS) {
  //check the settings here

  if (SETTINGS.N === "") {
    updateControlSheet("Please set a lookback window", SETTINGS);
  }
  if (fuzzyThresholdIsValid(SETTINGS['FUZZY_THRESHOLD']) === false) {
    updateControlSheet("Please check the fuzzy threshold", SETTINGS);
  }

  if (!sheetExists(SETTINGS)) {
    const errorMessage = `The sheet ${SETTINGS["TAB_NAME"]} does not exist. Please create it by duplicating an existing sheet (aka tab).`
    throw (errorMessage)
  }
}

function sheetExists(SETTINGS) {
  let sheets = SETTINGS.SPREADSHEET.getSheets();
  for (let sheet of sheets) {
    if (String(sheet.getName()) === String(SETTINGS["TAB_NAME"])) {
      return true;
    }
  }
  return false;
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

/*

    SET AND FORGET FUNCTIONS

    */

function getLogsColumn(controlSheet) {
  let col = 5;
  let LOGS_COLUMN = 0;
  while (String(controlSheet.getRange(3, col).getValue())) {
    LOGS_COLUMN = controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) {
      break;
    }
    col++;
  }
  return LOGS_COLUMN;
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
    AdWordsApp.currentAccount().getTimeZone(),
    format
  );
}

function log(msg) {
  try {
    console.log(AdWordsApp.currentAccount().getName() + " - " + msg);
  } catch (e) {
    log(msg);
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
    let ALL_SETTINGS = scanForAccounts();

    //   log(JSON.stringify(ALL_SETTINGS))

    //run all rows and all accounts
    for (let S in ALL_SETTINGS) {
      for (let R in ALL_SETTINGS[S]) {
        runScript(ALL_SETTINGS[S][R]);
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
    JSON.parse(INPUT)[AdWordsApp.currentAccount().getCustomerId().toString()];
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
    n = String(n).trim();
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
   * @param {number} - Cost
   * @returns {number}
   **/
  this.calculateRoas = function (ConversionValue, Cost) {
    if (Cost == 0) return 0;
    if (ConversionValue == 0) return 0;
    if (Cost > ConversionValue) return 0;
    return ConversionValue / Cost;
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
      AdWordsApp.currentAccount().getTimeZone(),
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
          AdWordsApp.currentAccount().getTimeZone(),
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

var FuzzySet = (function () {
  "use strict";

  const FuzzySet = function (
    arr,
    useLevenshtein,
    gramSizeLower,
    gramSizeUpper
  ) {
    var fuzzyset = {};

    // default options
    arr = arr || [];
    fuzzyset.gramSizeLower = gramSizeLower || 2;
    fuzzyset.gramSizeUpper = gramSizeUpper || 3;
    fuzzyset.useLevenshtein =
      typeof useLevenshtein !== "boolean" ? true : useLevenshtein;

    // define all the object functions and attributes
    fuzzyset.exactSet = {};
    fuzzyset.matchDict = {};
    fuzzyset.items = {};

    // helper functions
    var levenshtein = function (str1, str2) {
      var current = [],
        prev,
        value;

      for (var i = 0; i <= str2.length; i++)
        for (var j = 0; j <= str1.length; j++) {
          if (i && j)
            if (str1.charAt(j - 1) === str2.charAt(i - 1)) value = prev;
            else value = Math.min(current[j], current[j - 1], prev) + 1;
          else value = i + j;

          prev = current[j];
          current[j] = value;
        }

      return current.pop();
    };

    // return an edit distance from 0 to 1
    var _distance = function (str1, str2) {
      if (str1 === null && str2 === null)
        throw "Trying to compare two null values";
      if (str1 === null || str2 === null) return 0;
      str1 = String(str1);
      str2 = String(str2);

      var distance = levenshtein(str1, str2);
      if (str1.length > str2.length) {
        return 1 - distance / str1.length;
      } else {
        return 1 - distance / str2.length;
      }
    };

    // u00C0-u00FF is latin characters
    // u0621-u064a is arabic letters
    // u0660-u0669 is arabic numerals
    // TODO: figure out way to do this for more languages
    var _nonWordRe = /[^a-zA-Z0-9\u00C0-\u00FF\u0621-\u064A\u0660-\u0669, ]+/g;

    var _iterateGrams = function (value, gramSize) {
      gramSize = gramSize || 2;
      var simplified = "-" + value.toLowerCase().replace(_nonWordRe, "") + "-",
        lenDiff = gramSize - simplified.length,
        results = [];
      if (lenDiff > 0) {
        for (var i = 0; i < lenDiff; ++i) {
          simplified += "-";
        }
      }
      for (var i = 0; i < simplified.length - gramSize + 1; ++i) {
        results.push(simplified.slice(i, i + gramSize));
      }
      return results;
    };

    var _gramCounter = function (value, gramSize) {
      // return an object where key=gram, value=number of occurrences
      gramSize = gramSize || 2;
      var result = {},
        grams = _iterateGrams(value, gramSize),
        i = 0;
      for (i; i < grams.length; ++i) {
        if (grams[i] in result) {
          result[grams[i]] += 1;
        } else {
          result[grams[i]] = 1;
        }
      }
      return result;
    };

    // the main functions
    fuzzyset.get = function (value, defaultValue, minMatchScore) {
      // check for value in set, returning defaultValue or null if none found
      if (minMatchScore === undefined) {
        minMatchScore = 0.33;
      }
      var result = this._get(value, minMatchScore);
      if (!result && typeof defaultValue !== "undefined") {
        return defaultValue;
      }
      return result;
    };

    fuzzyset._get = function (value, minMatchScore) {
      var results = [];
      // start with high gram size and if there are no results, go to lower gram sizes
      for (
        var gramSize = this.gramSizeUpper;
        gramSize >= this.gramSizeLower;
        --gramSize
      ) {
        results = this.__get(value, gramSize, minMatchScore);
        if (results && results.length > 0) {
          return results;
        }
      }
      return null;
    };

    fuzzyset.__get = function (value, gramSize, minMatchScore) {
      var normalizedValue = this._normalizeStr(value),
        matches = {},
        gramCounts = _gramCounter(normalizedValue, gramSize),
        items = this.items[gramSize],
        sumOfSquareGramCounts = 0,
        gram,
        gramCount,
        i,
        index,
        otherGramCount;

      for (gram in gramCounts) {
        gramCount = gramCounts[gram];
        sumOfSquareGramCounts += Math.pow(gramCount, 2);
        if (gram in this.matchDict) {
          for (i = 0; i < this.matchDict[gram].length; ++i) {
            index = this.matchDict[gram][i][0];
            otherGramCount = this.matchDict[gram][i][1];
            if (index in matches) {
              matches[index] += gramCount * otherGramCount;
            } else {
              matches[index] = gramCount * otherGramCount;
            }
          }
        }
      }

      function isEmptyObject(obj) {
        for (var prop in obj) {
          if (obj.hasOwnProperty(prop)) return false;
        }
        return true;
      }

      if (isEmptyObject(matches)) {
        return null;
      }

      var vectorNormal = Math.sqrt(sumOfSquareGramCounts),
        results = [],
        matchScore;
      // build a results list of [score, str]
      for (var matchIndex in matches) {
        matchScore = matches[matchIndex];
        results.push([
          matchScore / (vectorNormal * items[matchIndex][0]),
          items[matchIndex][1],
        ]);
      }
      var sortDescending = function (a, b) {
        if (a[0] < b[0]) {
          return 1;
        } else if (a[0] > b[0]) {
          return -1;
        } else {
          return 0;
        }
      };
      results.sort(sortDescending);
      if (this.useLevenshtein) {
        var newResults = [],
          endIndex = Math.min(50, results.length);
        // truncate somewhat arbitrarily to 50
        for (var i = 0; i < endIndex; ++i) {
          newResults.push([
            _distance(results[i][1], normalizedValue),
            results[i][1],
          ]);
        }
        results = newResults;
        results.sort(sortDescending);
      }
      newResults = [];
      results.forEach(
        function (scoreWordPair) {
          if (scoreWordPair[0] >= minMatchScore) {
            newResults.push([
              scoreWordPair[0],
              this.exactSet[scoreWordPair[1]],
            ]);
          }
        }.bind(this)
      );
      return newResults;
    };

    fuzzyset.add = function (value) {
      var normalizedValue = this._normalizeStr(value);
      if (normalizedValue in this.exactSet) {
        return false;
      }

      var i = this.gramSizeLower;
      for (i; i < this.gramSizeUpper + 1; ++i) {
        this._add(value, i);
      }
    };

    fuzzyset._add = function (value, gramSize) {
      var normalizedValue = this._normalizeStr(value),
        items = this.items[gramSize] || [],
        index = items.length;

      items.push(0);
      var gramCounts = _gramCounter(normalizedValue, gramSize),
        sumOfSquareGramCounts = 0,
        gram,
        gramCount;
      for (gram in gramCounts) {
        gramCount = gramCounts[gram];
        sumOfSquareGramCounts += Math.pow(gramCount, 2);
        if (gram in this.matchDict) {
          this.matchDict[gram].push([index, gramCount]);
        } else {
          this.matchDict[gram] = [[index, gramCount]];
        }
      }
      var vectorNormal = Math.sqrt(sumOfSquareGramCounts);
      items[index] = [vectorNormal, normalizedValue];
      this.items[gramSize] = items;
      this.exactSet[normalizedValue] = value;
    };

    fuzzyset._normalizeStr = function (str) {
      if (Object.prototype.toString.call(str) !== "[object String]")
        throw "Must use a string as argument to FuzzySet functions";
      return str.toLowerCase();
    };

    // return length of items in set
    fuzzyset.length = function () {
      var count = 0,
        prop;
      for (prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          count += 1;
        }
      }
      return count;
    };

    // return is set is empty
    fuzzyset.isEmpty = function () {
      for (var prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          return false;
        }
      }
      return true;
    };

    // return list of values loaded into set
    fuzzyset.values = function () {
      var values = [],
        prop;
      for (prop in this.exactSet) {
        if (this.exactSet.hasOwnProperty(prop)) {
          values.push(this.exactSet[prop]);
        }
      }
      return values;
    };

    // initialization
    var i = fuzzyset.gramSizeLower;
    for (i; i < fuzzyset.gramSizeUpper + 1; ++i) {
      fuzzyset.items[i] = [];
    }
    // add all the items to the set
    for (i = 0; i < arr.length; ++i) {
      fuzzyset.add(arr[i]);
    }

    return fuzzyset;
  };

  return FuzzySet;
})();

function fuzzyStringInArray(string, array, fuzzyThreshold) {
  for (let i in array) {
    let keyword = array[i]
    if (fuzzyMatchScore(string, keyword) > fuzzyThreshold) {
      // log(`Found match: ${string} and ${keyword} with score ${fuzzyMatchScore(string, keyword)}`)
      return true
    }
  }
  return false
}

function fuzzyMatchScore(needle, haystack) {
  let a = FuzzySet();
  a.add(haystack);
  let result = a.get(needle)
  if (!result) return 0
  return result[0][0];
}

/**
 * get keywords that are similar to the search term
 * @param {string} searchTerm - the search term
 * @param {array} keywords - the keywords to search through
 * @returns {array} - the keywords that are similar to the search term
 */
function getSimilarKeywords(searchTerm, keywords, fuzzyMatchMoreThan, fuzzyMatchLessThan) {
  let checkedKeywords = []
  let similarKeywords = []
  for (let i in keywords) {
    let k = keywords[i]
    //skip k if it's already in checkedKeywords
    if (checkedKeywords.indexOf(k) > -1) {
      continue;
    }
    checkedKeywords.push(k)

    let score = fuzzyMatchScore(searchTerm, k)
    if (score > fuzzyMatchMoreThan && score <= fuzzyMatchLessThan) {
      let roundedScore = roundFloatToTwoDecimalPlaces(score)
      similarKeywords.push(`${k} (${roundedScore})`)
    }
  }
  return similarKeywords
}

function roundFloatToTwoDecimalPlaces(float) {
  return Math.round(float * 100) / 100
}

/**
 * validate the theshold setting
 * @param {number} theshold - the threshold
 * @returns {bool} - whether the threshold is valid
 */
function fuzzyThresholdIsValid(theshold) {
  if (theshold > 1) {
    return false
  }
  if (theshold < 0) {
    return false
  }
  if (String(theshold) === "") {
    return false
  }
  return true
}