/**
 * Negative Keyword Finder
 * @Author Charles Bannister (shabba.io)
 * @Version 1.3
 * View the script at https://shabba.io/script/7
**/

// Template: https://docs.google.com/spreadsheets/d/1MvCwzNCOIG3AO34b5yVe5VjcNFDwdvhCJdhrZQRXz58/edit?gid=528576320#gid=528576320
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";


//TODO
// Contains / not contains (update template too) - done
// Auto-create output sheets (name it based on the rule number)
// Slack
// Notification frequency
// Subject prefix
// List negatives (grab from the main settings)
// Manager version
// Master filters?



const INPUT_TAB_NAME = "Settings";

//set to true to send emails regardless of preview status
const EMAIL_DURING_PREVIEW = false;

// Change this number if more filters are added
const NUMBER_OF_FILTERS = 6;



const SCRIPT_NAME = "Negative Keyword Finder"

const SHABBA_SCRIPT_ID = 7;

const VERSION = "1.3";



//No need to edit anything below this line

const h = new Helper();
const s = new Setting();

const OPTIONS = { includeZeroImpressions: true };

const SINGLE_ACCOUNT_HEADER_TYPES = {
  NAME: "normal",
  EMAILS: "csv",
  FLAG: "bool",
  EMAIL_ALERT: "bool",
  CHANGES_EMAIL: "bool",
  TAB_NAME: "normal",
  N: "normal",
  CAMPAIGN_NAME_CONTAINS: "csv",
  CAMPAIGN_NAME_NOT_CONTAINS: "csv",
  QUERY_TARGETING_STATUS: "normal",
  DEFAULT_MATCH_TYPE: "normal",
  DEFAULT_ACTION: "normal",
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
  console.log(JSON.stringify(SETTINGS));

  SETTINGS.SPREADSHEET = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
  SETTINGS.CONTROL_SHEET = SETTINGS.SPREADSHEET.getSheetByName(INPUT_TAB_NAME);

  checkSettings(SETTINGS);

  processRowSettings(SETTINGS);

  addLogSheetInfo(SETTINGS);

  log(`Settings: ${JSON.stringify(SETTINGS)}`);

  const queries = getQueriesFromSheet(SETTINGS);

  // log(JSON.stringify(queries))


  addKeywords(SETTINGS, queries);

  //check queries and add them to the sheet
  addQueriesToSheet(SETTINGS);

  SETTINGS.LOGS.push("The script ran successfully");
  updateControlSheet("", SETTINGS);

  log("Finished");
}

function getQueriesFromSheet(SETTINGS) {
  var map = { adgroup_ids: [], rows: {} };

  var sheet = SETTINGS.KEYWORD_LOG_SHEET;

  var data = sheet.getDataRange().getValues();
  data.shift();
  data.shift();
  var header = data.shift();

  for (var d in data) {
    var row = data[d];

    var query = String(row[header.indexOf("Search Term")]);

    var addToAdGroup = row[header.indexOf("Add as Keyword")];

    var match_type = row[header.indexOf("Match Type")];

    var this_is_ok = row[header.indexOf("Ignore")];

    var add_as_negative = row[header.indexOf("Add as Ad Group Negative")];

    if (!addToAdGroup && !add_as_negative & !this_is_ok) continue;

    var adgroup_id = row[header.indexOf("Ad Group Id")];

    var campaignName = row[header.indexOf("Campaign Name")];

    var adGroupName = row[header.indexOf("Adgroup")];

    var averageCpc =
      parseFloat(row[header.indexOf("Cost")]) /
      parseFloat(row[header.indexOf("Clicks")]);

    var clicks = row[header.indexOf("Clicks")];

    var cost = row[header.indexOf("Cost")];

    var conversions = row[header.indexOf("Conversions")];

    var cpa = row[header.indexOf("Cost / Conv.")];

    var impressions = row[header.indexOf("Impr.")];

    var conversion_rate = row[header.indexOf("Conv. Rate")];

    if (map["adgroup_ids"].indexOf(adgroup_id) === -1)
      map["adgroup_ids"].push(adgroup_id);

    map["rows"][adgroup_id] = map["rows"][adgroup_id] || {};
    map["rows"][adgroup_id][query] = {};
    map["rows"][adgroup_id][query]["query"] = query;
    map["rows"][adgroup_id][query]["add_to_adgroup"] = addToAdGroup;
    map["rows"][adgroup_id][query]["add_as_negative"] = add_as_negative;
    map["rows"][adgroup_id][query]["match_type"] = match_type;
    map["rows"][adgroup_id][query]["campaignName"] = campaignName;
    map["rows"][adgroup_id][query]["adGroupName"] = adGroupName;
    map["rows"][adgroup_id][query]["adGroupId"] = adgroup_id;
    map["rows"][adgroup_id][query]["averageCpc"] = averageCpc;
    map["rows"][adgroup_id][query]["clicks"] = clicks;
    map["rows"][adgroup_id][query]["cost"] = cost;
    map["rows"][adgroup_id][query]["conversions"] = conversions;
    map["rows"][adgroup_id][query]["cpa"] = cpa;
    map["rows"][adgroup_id][query]["impressions"] = impressions;
    map["rows"][adgroup_id][query]["conversion_rate"] = conversion_rate;
    map["rows"][adgroup_id][query]["this_is_ok"] = this_is_ok;
  }

  return map;
}

function addLogSheetInfo(SETTINGS) {
  SETTINGS.LOG_SHEET_URL = INPUT_SHEET_URL;

  var logSS = SpreadsheetApp.openByUrl(SETTINGS.LOG_SHEET_URL);
  SETTINGS.logSS = logSS;
  SETTINGS.KEYWORD_LOG_SHEET = logSS.getSheetByName(SETTINGS["TAB_NAME"]);
}

function sendChangesEmail(SETTINGS) {
  if (SETTINGS.PREVIEW_MODE && !EMAIL_DURING_PREVIEW) return;

  //Send email
  var SUB =
    AdsApp.currentAccount().getName() + " - " + SCRIPT_NAME + " script.";
  var MSG =
    "Hi,<br><br>The " +
    INPUT_TAB_NAME +
    " script ran successfully and changes were made. Here are the logs:<br><br>";

  MSG += "<ul>";
  for (var l in SETTINGS.LOGS) {
    MSG += "<li>";
    MSG += SETTINGS.LOGS[l];
    MSG += "</li>";
  }
  MSG += "</ul>";

  MSG +=
    "<br><br>Here's the sheet where you'll find the settings and Search Term data:<br>" +
    SETTINGS.LOG_SHEET_URL;
  MSG += "<br><br>Thanks,";
  MSG += "<br><br>Charles";
  MSG += "<br><br><br><a href='https://shabba.io'>shabba.io</a>";
  var emails = SETTINGS.EMAILS;

  for (var i in emails) {
    MailApp.sendEmail({
      to: emails[i],
      subject: SUB,
      htmlBody: MSG,
    });
  }
}

function sendEmailAlert(SETTINGS, numberOfSearchTerms) {
  if (SETTINGS.PREVIEW_MODE && !EMAIL_DURING_PREVIEW) return;

  SETTINGS.LOGS.push(`Rule name: ${SETTINGS["NAME"]}`);
  SETTINGS.LOGS.push(`${numberOfSearchTerms} search terms matched the rule`);

  //Send email
  var SUB = AdsApp.currentAccount().getName() + " - Search Terms Alert";
  var MSG =
    "Hi,<br><br>The Search Terms Manager script has results. Here are the logs:<br><br>";

  MSG += "<ul>";
  for (var l in SETTINGS.LOGS) {
    MSG += "<li>";
    MSG += SETTINGS.LOGS[l];
    MSG += "</li>";
  }
  MSG += "</ul>";

  MSG +=
    "<br><br>Visit the Google Sheet to review the terms:<br>" +
    SETTINGS.LOG_SHEET_URL;
  MSG += "<br><br>Thanks,";
  MSG += "<br><br><br><a href='https://shabba.io'>shabba.io</a>";
  var emails = SETTINGS.EMAILS;

  for (var i in emails) {
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
  var put = [logString, SETTINGS.NOW];
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
  for (var key in SETTINGS) {
    if (key.indexOf("FILTER") === -1 || key.indexOf("METRIC") === -1) {
      continue;
    }
    if (SETTINGS[key] !== "Labels") continue;
    log(key + " - " + SETTINGS[key]);
    var filter_value = SETTINGS[key.replace("METRIC", "VALUE")];
    var value_split = String(filter_value).split(",");
    for (var v in value_split) {
      value_split[v] = getLabelIdFromString(value_split[v]);
    }
    SETTINGS[key.replace("METRIC", "VALUE")] = value_split.join(",");
  }
  return SETTINGS;
}

function getFilterWhereString(SETTINGS) {
  //custom metrics map
  //name : formula, whether high or low is good (e.g. ROAS high is good, CPA high is bad - used for Vs target calcs)

  var filters = Object.keys(SETTINGS).map(function (x) {
    if (x.indexOf("FILTER") > -1) {
      return x;
    }
  });

  var filterMap = {};
  var filterParts = ["metric", "operator", "value"];
  var numberOfFilters = filters.length / filterParts.length;

  // log("number of filters: " + numberOfFilters)
  for (var i = 0; i < numberOfFilters; i++) {
    var filterName = "FILTER_" + (i + 1);
    if (Object.keys(SETTINGS).indexOf(filterName + "_METRIC") == -1) continue;
    if (SETTINGS[filterName + "_METRIC"] == "") continue;
    filterMap[filterName] = filterMap[filterName] || {};

    for (var x in filterParts) {
      filterMap[filterName][filterParts[x]] =
        SETTINGS[filterName + "_" + filterParts[x].toUpperCase()];
    }
  }

  var where = "";

  var whereArray = [];

  //turn the filters object into a where statement string
  whereArray.push(filtersToWhereStatement(filterMap, filterParts));
  function filtersToWhereStatement(filterMap, filterParts) {
    var str = [];
    for (var filter in filterMap) {
      //if the metric is a custom metric, continue
      if (Object.keys(CUSTOM_METRICS).indexOf(filterMap[filter]["metric"]) > -1)
        continue;

      str.push("and");
      for (var p in filterParts) {
        str.push(filterMap[filter][filterParts[p]]);
      }
    }
    return str.join(" ");
  }

  where += whereArray.join(" ");

  return where;
}

function getLabelIdFromString(str) {
  var labels = AdsApp.labels()
    .withCondition("Name = '" + str.trim() + "'")
    .get();
  if (labels.hasNext()) {
    var label = labels.next();
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

  var labels = AdsApp.labels().get();
  var exists = false;
  while (labels.hasNext()) {
    var label = labels.next();
    if (label.getName() == labelName) {
      exists = true;
    }
  }
  if (!exists) {
    AdsApp.createLabel(labelName);
  }
}

function getQueryWhereString(SETTINGS) {
  var where = "where CampaignStatus = ENABLED and AdGroupStatus = ENABLED ";
  if (SETTINGS.QUERY_TARGETING_STATUS) {
    where += " and QueryTargetingStatus = " + SETTINGS.QUERY_TARGETING_STATUS;
  }

  var whereArray = [];

  for (var i in SETTINGS.CAMPAIGN_NAME_CONTAINS) {
    whereArray.push(
      " and CampaignName CONTAINS_IGNORE_CASE '" +
      SETTINGS.CAMPAIGN_NAME_CONTAINS[i].trim() +
      "'"
    );
  }

  for (var i in SETTINGS.CAMPAIGN_NAME_NOT_CONTAINS) {
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
  for (var metricName in CUSTOM_METRICS) {
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

function getKeywords(SETTINGS, sheet_name) {
  let sheetOperations = new SheetOperations(SETTINGS);
  var sheet = sheetOperations.getSheetByName(sheet_name);

  var data = sheetOperations.getAllRowsData(sheet, 5, 3);
  var query_keys = [];
  for (var d in data) {
    var row = data[d];
    if (row[0] === "") break;
    query_keys.push(row.join(""));
  }

  return query_keys;
}
function getPositiveKeywordsQueries(SETTINGS) {
  return getKeywords(SETTINGS, "Ignored Terms");
}

function getNegativeKeywordsQueries(SETTINGS) {
  return getKeywords(SETTINGS, "Negative Keywords");
}

function addQueriesToSheet(SETTINGS) {
  log(`Clearing the ${SETTINGS["TAB_NAME"]} sheet`);
  let sheetOperations = new SheetOperations(SETTINGS);
  var sheet = sheetOperations.getSheetByName(SETTINGS["TAB_NAME"]);
  if (sheet.getMaxRows() > 3) {
    const rangeToClear = sheet.getRange(4, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    rangeToClear.clear();
  }

  sheet.setFrozenRows(3);

  SETTINGS["FILTER_MAP"] = getFilterMap(SETTINGS);

  var positive_keywords_query_keys = getPositiveKeywordsQueries(SETTINGS);
  var negative_keywords_query_keys = getNegativeKeywordsQueries(SETTINGS);

  var cols = [
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

  var reportName = "SEARCH_QUERY_PERFORMANCE_REPORT";

  var query = [
    "select",
    cols.join(","),
    "from",
    reportName,
    getQueryWhereString(SETTINGS),
    "during",
    SETTINGS.DATE_RANGE,
  ].join(" ");

  log(reportName + " query: \n");
  console.log(query.split('from')[0], '\n from', query.split('from')[1], '\n');

  var logArray = [];

  var reportIter = AdsApp.report(query, OPTIONS).rows();
  while (reportIter.hasNext()) {
    var row = reportIter.next();

    row.Impressions = parseInt(row.Impressions, 10);
    row.Clicks = parseInt(row.Clicks, 10);
    row.Conversions = parseFloat(row.Conversions.toString().replace(/,/g, ""));

    row.Cost = parseFloat(row.Cost.toString().replace(/,/g, ""));
    row.ConversionValue = parseFloat(
      row.ConversionValue.toString().replace(/,/g, "")
    );

    var query_key = row.CampaignName + row.AdGroupName + row.Query;
    if (positive_keywords_query_keys.indexOf(query_key) > -1) {
      continue;
    }
    if (negative_keywords_query_keys.indexOf(query_key) > -1) {
      continue;
    }

    row = addCustomMetricsToRow(row);

    if (skipEntity(row, SETTINGS)) continue;
    let addToAdgroup =
      SETTINGS.DEFAULT_ACTION.trim() === "Add Keyword" ? "TRUE" : "FALSE";
    let addAsNegative =
      SETTINGS.DEFAULT_ACTION.trim() === "Add Negative Keyword"
        ? "TRUE"
        : "FALSE";
    var logRow = [
      row.CampaignName,
      row.AdGroupName,
      row.AdGroupId,
      row.Query,
      row.QueryTargetingStatus,
      addToAdgroup,
      SETTINGS["DEFAULT_MATCH_TYPE"],
      false,
      addAsNegative,
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

    // if(logArray.length > 7)break
  }

  log(logArray.length + " search terms found");
  if (logArray.length === 0) return;

  sheetOperations.writeToSheet(logArray, SETTINGS["TAB_NAME"], 4);

  postWriteOperations(logArray, SETTINGS);

  if (SETTINGS["EMAIL_ALERT"]) {
    sendEmailAlert(SETTINGS, logArray.length);
  }
}

function postWriteOperations(logArray, SETTINGS) {
  var sheetOperations = new SheetOperations(SETTINGS);
  var sheet = sheetOperations.getSheetByName(SETTINGS["TAB_NAME"]);
  sheet
    .getRange("A2")
    .setValue(
      "Current Data For Lookback (" +
      SETTINGS.N +
      " days)" +
      " - " +
      SETTINGS.DATE_RANGE
    );
  var header = sheet.getRange(3, 1, 1, sheet.getLastColumn()).getValues();
  addValidation(logArray.length, sheet, header, sheetOperations);
  sheetOperations.sortColumn(sheet, parseInt(header[0].indexOf("Cost")) + 1);
}

function addValidation(rows, sheet, header, sheetOperations) {
  //add validaiton to the keywords sheet
  var checkbox_cols = [];
  var matchTypeColumn = 0;
  for (var value_index in header[0]) {
    var value = header[0][value_index];
    if (value == "Add as Keyword")
      checkbox_cols.push(parseInt(value_index) + 1);
    if (value == "Ignore") checkbox_cols.push(parseInt(value_index) + 1);
    if (value == "Add as Ad Group Negative") {
      checkbox_cols.push(parseInt(value_index) + 1);
    }
    if (value == "Match Type") {
      matchTypeColumn = parseInt(value_index) + 1;
    }
  }
  if (matchTypeColumn === 0) {
    throw new Error(`There was a problem finding the "Match Type" column in the ${sheet.getName()} sheet. 
    Please make sure the column exists and try again.
    It might help to grab the original template sheet from https://shabba.io/script/${SHABBA_SCRIPT_ID}`);
  }
  var start_row = 4;
  for (var checkbox_cols_index in checkbox_cols) {
    sheetOperations.addCheckBox(
      start_row,
      checkbox_cols[checkbox_cols_index],
      rows,
      sheet
    );
  }
  // log('These columns should match the keywords sheet:')
  // log('Match type column: ' + matchTypeColumn)
  // log('Checkbox columns: ' + checkbox_cols)
  var validation_list = ["Broad", "Phrase", "Exact"];
  sheetOperations.addDropdownValidation(
    validation_list,
    sheet,
    start_row,
    rows,
    matchTypeColumn
  );
}

class SheetOperations {
  constructor(SETTINGS) {
    this.SETTINGS = SETTINGS;
  }

  addDropdownValidation(validation_list, sheet, start_row, rows, column) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(validation_list)
      .build();
    sheet.getRange(start_row, column, rows, 1).setDataValidation(rule);
  }

  addCheckBox(start_row, column, rows, sheet) {
    var range = sheet.getRange(start_row, column, rows, 1);

    var enforceCheckbox = SpreadsheetApp.newDataValidation();
    enforceCheckbox.requireCheckbox();
    enforceCheckbox.setAllowInvalid(false);
    enforceCheckbox.build();

    range.setDataValidation(enforceCheckbox);
  }

  getSheetByName(name) {
    try {
      return this.SETTINGS.logSS.getSheetByName(name);
    } catch (e) {
      var sheets = this.SETTINGS.logSS.getSheets();
      var sheet_names = [];
      for (var sheet_key in sheets) {
        var sheet = sheets[sheet_key];
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

    var sheet = this.getSheetByName(sheetName);

    sheet.insertRowsAfter(start_row - 1, logArray.length);

    var max_rows = 20000;

    if (sheet.getLastRow() > max_rows) {
      sheet.deleteRows(max_rows, sheet.getLastRow() - max_rows);
    }

    sheet
      .getRange(start_row, 1, logArray.length, logArray[0].length)
      .setValues(logArray);
  }

  getAllRowsData(sheet, start_row, num_columns) {
    var start_column = 1;
    return sheet
      .getRange(start_row, start_column, sheet.getLastRow(), num_columns)
      .getValues();
  }

  writeToSheet(logArray, sheetName, start_row) {
    log("Writing to " + sheetName);
    if (logArray.length === 0) return;

    var sheet = this.getSheetByName(sheetName);
    var lastRow = this.getLastRow(sheet, 4);
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
    var values = [];
    for (var r = 0; r < rows; r++) {
      var this_row = [];
      for (var c = 0; c < columns; c++) {
        this_row.push("");
      }
      values.push(this_row);
    }
    sheet.getRange(row, column, rows, columns).setValues(values);
  }

  /*Get the last row based on the first column's values */
  getLastRow(sheet, startRow) {
    var row = startRow;
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
    const startRow = 4;
    const numberOfRows = sheet.getLastRow() - startRow
    if (numberOfRows < 2) {
      return
    }
    sheet
      .getRange(
        startRow,
        1,
        numberOfRows,
        sheet.getLastColumn()
      )
      .sort({ column: columnNumber, ascending: false });
  }
}

function getFilterMap(SETTINGS) {
  var filters = Object.keys(SETTINGS).map(function (x) {
    if (x.indexOf("FILTER") > -1) {
      return x;
    }
  });

  var filterMap = {};
  var filterParts = ["metric", "operator", "value"];
  var numberOfFilters = filters.length / filterParts.length;
  // log("number of filters: " + numberOfFilters)
  for (var i = 0; i < numberOfFilters; i++) {
    var filterName = "FILTER_" + (i + 1);
    if (Object.keys(SETTINGS).indexOf(filterName + "_METRIC") == -1) continue;
    SETTINGS[filterName + "_METRIC"] = renameMetricToMatchApi(SETTINGS[filterName + "_METRIC"]);
    if (SETTINGS[filterName + "_METRIC"] == "") {
      continue;
    }
    filterMap[filterName] = filterMap[filterName] || {};
    for (var x in filterParts) {
      filterMap[filterName][filterParts[x]] =
        SETTINGS[filterName + "_" + filterParts[x].toUpperCase()];
    }
  }

  return filterMap;
}


/**
 * Where a user friendly metric/attribute name is used,
 * this function renames it to the API name
 */
function renameMetricToMatchApi(metric) {
  const metricMap = {
    "SearchTerm": "Query",
  }
  return metricMap[metric] || metric;
}

//Whether to skip the entity based on the stats and filters
//Only check custom metrics e.g. cpa
function skipEntity(row, SETTINGS) {
  //custom metrics map
  //name : formula, whether high or low is good (e.g. ROAS high is good, CPA high is bad - used for Vs target calcs)
  var filterMap = SETTINGS["FILTER_MAP"];

  // log(JSON.stringify(filterMap))

  // log(JSON.stringify(row))

  for (var filter in filterMap) {
    var this_filter = filterMap[filter];
    if (filterNotInCustomMetrics(this_filter.metric)) continue;
    var eval_string =
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
  var result = false;

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
function addKeywords(SETTINGS, queries) {
  let adGroupIds = queries["adgroup_ids"];

  if (adGroupIds.length === 0) {
    log("No negative keywords to add");
    return;
  }

  const negativeKeywordLogArray = [];
  const positiveKeywordLogArray = [];

  const chunkSize = 10000;

  const types = ["search", "shopping"];

  for (const typeNum in types) {
    const chunkedArray = [];
    for (var i = 0; i < adGroupIds.length; i += chunkSize) {
      chunkedArray.push(adGroupIds.slice(i, i + chunkSize));
    }

    for (let i = 0; i < chunkedArray.length; i++) {
      let adGroups = types[typeNum] === "shopping" ? AdsApp.shoppingAdGroups() : AdsApp.adGroups();
      adGroups = adGroups
        .withIds(chunkedArray[i])
        .get();

      if (!adGroups.hasNext()) {
        continue;
      }

      while (adGroups.hasNext()) {
        const adGroup = adGroups.next();
        const adGroupId = adGroup.getId();

        const adGroupQueries = queries["rows"][adGroupId];

        var previewModeText = SETTINGS.PREVIEW_MODE ? " (Preview Mode) " : "";

        for (const query in adGroupQueries) {
          const row = adGroupQueries[query];

          if (row["this_is_ok"]) {
            const logRow = [
              row["campaignName"],
              row["adGroupName"],
              query,
              row["match_type"],
              SETTINGS.NOW,
              previewModeText,
            ];

            positiveKeywordLogArray.push(logRow);

            continue;
          }

          if (row["add_as_negative"]) {
            const keyword = addMatchType(query, row["match_type"]);
            adGroup.createNegativeKeyword(keyword);
            const logRow = [
              row["campaignName"],
              row["adGroupName"],
              query,
              row["match_type"],
              SETTINGS.NOW,
              previewModeText,
            ];
            negativeKeywordLogArray.push(logRow);
            continue;
          }
          if (row["add_to_adgroup"]) {
            const keyword = addMatchType(query, row["match_type"]);
            adGroup
              .newKeywordBuilder()
              .withText(keyword)
              .withCpc(row["cost"] / row["clicks"])
              .build();
          }
        }
      }
    }
  }

  const sheetOperations = new SheetOperations(SETTINGS);
  sheetOperations.appendToSheet(
    negativeKeywordLogArray,
    "Negative Keywords",
    5
  );
  SETTINGS.LOGS.push(
    parseInt(negativeKeywordLogArray.length) + " negative keywords added"
  );

  sheetOperations.appendToSheet(positiveKeywordLogArray, "Ignored Terms", 5);
  SETTINGS.LOGS.push(
    parseInt(positiveKeywordLogArray.length) + " positive keywords added"
  );

  if ((positiveKeywordLogArray.length > 0 || negativeKeywordLogArray.length > 0) && SETTINGS['CHANGES_EMAIL']) {
    sendChangesEmail(SETTINGS);
  }
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

  var map = {};

  var sheet = SETTINGS.CHANGE_LOG_SHEET;
  if (sheet.getLastRow() < 2) {
    log("No change log values");
    map = updateMap(map, CampaignName, AdGroupName, Id, SETTINGS.NOW);
    writeMapToSheet(sheet, map);
    return true;
  }

  //grab the current details and update
  var data = sheet.getDataRange().getValues();
  data.shift();
  for (var d in data) {
    var row = data[d];
    var campaign = row[0];
    var adgroup = row[1];
    var id = row[2];
    var date = row[3];
    map[campaign + adgroup + id] = {};
    map[campaign + adgroup + id]["campaign"] = campaign;
    map[campaign + adgroup + id]["adgroup"] = adgroup;
    map[campaign + adgroup + id]["id"] = id;
    map[campaign + adgroup + id]["date"] = date;
  }

  var this_id = CampaignName + AdGroupName + Id;

  if (valueUndefined(map[this_id])) {
    // log("Change log id undefined: " + this_id)
    // log(JSON.stringify(map))
    map = updateMap(map, CampaignName, AdGroupName, Id, SETTINGS.NOW);
    writeMapToSheet(sheet, map);
    return true;
  }

  var this_date = new Date(map[this_id]["date"]);

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
    var header = ["Campaign", "Ad group", "Id", "Last Updated"];
    var logArray = [header];
    for (var id in map) {
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
    var this_id = CampaignName + AdGroupName + Id;
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
  var YESTERDAY = getAdWordsFormattedDate(1, "yyyyMMdd");

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
  var editors = CONTROL_SHEET.getRange(1, logsColumn).getValue();
  if (editors == "") {
    return;
  }
  if (editors.indexOf(",") > -1) {
    editors = editors.split(",");
    for (var e in editors) {
      editors[e] = editors[e].trim().toLowerCase();
    }
  } else {
    editors = [editors.trim().toLowerCase()];
  }
  return editors;
}

function getFilterHeaderTypes() {
  var map = {};
  for (var i = 1; i < NUMBER_OF_FILTERS + 1; i++) {
    map["FILTER_" + i + "_METRIC"] = "normal";
    map["FILTER_" + i + "_OPERATOR"] = "filter_operator";
    map["FILTER_" + i + "_VALUE"] = "filter_value";
  }
  return map;
}

function getHeaderTypes() {
  const MCC_HEADER_TYPES = {
    'ACCOUNT_ID': 'normal',
    'ACCOUNT_NAME': 'normal',
  };

  const FILTER_HEADER_TYPES = getFilterHeaderTypes();

  let HEADER_TYPES = {};

  if (isMCC()) {
    HEADER_TYPES = objectMerge(
      MCC_HEADER_TYPES,
      SINGLE_ACCOUNT_HEADER_TYPES,
      FILTER_HEADER_TYPES,
      SINGLE_ACCOUNT_HEADER_TYPES_AFTER
    );
  } else {
    HEADER_TYPES = objectMerge(
      SINGLE_ACCOUNT_HEADER_TYPES,
      FILTER_HEADER_TYPES,
      SINGLE_ACCOUNT_HEADER_TYPES_AFTER
    );
  }
  return HEADER_TYPES;
}

function buildTestSettings() {
  var map = {};
  for (var key in SINGLE_ACCOUNT_HEADER_TYPES) {
    map[key] = "random value";
  }
  return map;
}

function scanForAccounts() {
  log("getting settings...");
  log(
    "The settings sheet should contain " + NUMBER_OF_FILTERS + " filter sets"
  );
  var map = {};
  var controlSheet =
    SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME);
  var data = controlSheet.getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();

  const HEADER_TYPES = getHeaderTypes();
  // console.log(`HEADER_TYPES: ${JSON.stringify(HEADER_TYPES)}`);

  const HEADER = Object.keys(HEADER_TYPES);
  // console.log(`HEADER: ${HEADER}`);

  var LOGS_COLUMN = 0;
  var col = 5;
  while (controlSheet.getRange(3, col).getValue()) {
    LOGS_COLUMN = controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) {
      break;
    }
    col++;
  }
  // console.log(`LOGS_COLUMN: ${LOGS_COLUMN}`);

  // log(HEADER)
  const flagIndex = 2;
  const ruleNameIndex = 0;

  for (const row of data) {
    //if "run script" is not set to "yes", continue.
    // log(data[k][flagIndex])
    const rowNum = data.indexOf(row) + 4;
    if (row[ruleNameIndex] === "") {
      console.log(`Rule name is empty for row ${rowNum}. Skipping...`);
      continue;
    }
    if (row[flagIndex].toLowerCase() != "yes") {
      console.log(`Flag set to ${row[flagIndex]} for row ${rowNum}. Skipping...`);
      continue;
    }

    const id = row[0];
    const rowId = id + "/" + rowNum;
    map[id] = map[id] || {};
    map[id][rowId] = { ROW_NUM: rowNum };
    for (const key in HEADER) {
      if (HEADER[key] == "LOGS_COLUMN") {
        map[id][rowId][HEADER[key]] = LOGS_COLUMN;
        continue;
      }
      map[id][rowId][HEADER[key]] = row[key];
    }
  }

  if (Object.keys(map).length == 0) {
    console.warn("No rules were specified");
    console.warn(
      "To run the script, add settings and set 'Run Script' to 'Yes'"
    );
    return;
  }

  var previousOperator = "";
  for (var id in map) {
    for (var rowId in map[id]) {
      for (var key in map[id][rowId]) {
        var isFilterValue =
          key.indexOf("FILTER") > -1 && key.indexOf("VALUE") > -1;
        var filter_metric = !isFilterValue
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
  for (var i = 1; i < arguments.length; i++)
    for (var a in arguments[i]) arguments[0][a] = arguments[i][a];
  return arguments[0];
}

function isList(operator) {
  var list_operators = [
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
  var arr = String(value).split(",");
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
  var type = HEADER[key];
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
      var ret = value.split(",");
      ret = ret[0] == "" && ret.length == 1 ? [] : ret;
      if (ret.length == 0) {
        return [];
      } else {
        for (var r in ret) {
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
    AdsApp.currentAccount().getTimeZone(),
    "MMM dd, yyyy HH:mm:ss"
  );
  SETTINGS["QUERY_TARGETING_STATUS"] =
    SETTINGS["QUERY_TARGETING_STATUS"] == "ALL"
      ? ""
      : SETTINGS["QUERY_TARGETING_STATUS"];

  SETTINGS.LOGS_COLUMN = getLogsColumn(SETTINGS.CONTROL_SHEET);
  SETTINGS.LOGS = [];
  SETTINGS.EMAILS_SHEET = true;

  var defaultNote =
    "Possible problems include: 1) There was an error (check the logs within Google Ads) 2) The script was stopped before completion";
  SETTINGS.CONTROL_SHEET.getRange(SETTINGS.ROW_NUM, SETTINGS.LOGS_COLUMN, 1, 1)
    .setValue(
      "The script is either still running or didn't finish successfully"
    )
    .setNote(defaultNote);

  parseDateRange(SETTINGS);
  SETTINGS.PREVIEW_MODE = AdsApp.getExecutionInfo().isPreview();

  if (SETTINGS.PREVIEW_MODE) {
    var msg = "Running in preview mode. No changes will be made.";
    SETTINGS.LOGS.push(msg);
  }

  //  checkLabel(SETTINGS.KEYWORD_LABEL)
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
  var currentEditors = spreadsheet.getEditors();
  var currentEditorEmails = [];
  for (var c in currentEditors) {
    currentEditorEmails.push(currentEditors[c].getEmail().trim().toLowerCase());
  }

  var editors = controlSheet.getRange(1, LOGS_COLUMN).getValue();
  if (editors == "") {
    return;
  }
  if (editors.indexOf(",") > -1) {
    editors = editors.split(",");
    for (var e in editors) {
      editors[e] = editors[e].trim().toLowerCase();
    }
  } else {
    editors = [editors.trim().toLowerCase()];
  }

  for (var e in editors) {
    var index = currentEditorEmails.indexOf(editors[e]);
    if (currentEditorEmails.indexOf(editors[e]) == -1) {
      spreadsheet.addEditor(editors[e]);
    }
  }
}

/*
 
    SET AND FORGET FUNCTIONS
 
    */

function getLogsColumn(controlSheet) {
  var col = 5;
  var LOGS_COLUMN = 0;
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
  var s = "";
  for (var l in logs) {
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
  var date = new Date();
  date.setDate(date.getDate() - d);
  return Utilities.formatDate(
    date,
    AdsApp.currentAccount().getTimeZone(),
    format
  );
}

function log(msg) {
  try {
    console.log(AdsApp.currentAccount().getName() + " - " + msg);
  } catch (e) {
    log(msg);
  }
}

function round(num, n) {
  return +(Math.round(parseFloat(num) + "e+" + n) + "e-" + n);
}

function main() {

  runTopLevelLogic();
}


function runTopLevelLogic() {
  if (isMCC()) {
    var SETTINGS = scanForAccounts();
    log(JSON.stringify(SETTINGS))
    var ids = Object.keys(SETTINGS);
    if (ids.length == 0) {
      Logger.log("No Rules Specified");
      return;
    }
    MccApp.accounts()
      .withIds(ids)
      .withLimit(50)
      .executeInParallel("runRows", "callBack", JSON.stringify(SETTINGS));
  } else {
    var ALL_SETTINGS = scanForAccounts();

    log(JSON.stringify(ALL_SETTINGS))

    //run all rows and all accounts
    for (var S in ALL_SETTINGS) {
      for (var R in ALL_SETTINGS[S]) {
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
  var SETTINGS =
    JSON.parse(INPUT)[AdsApp.currentAccount().getCustomerId().toString()];
  for (var rowId in SETTINGS) {
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
    var digits = n.split("");
    for (var d in digits) {
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
    var col = 5;
    var LOGS_COLUMN = 0;
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
    var s = "";
    for (var l in logs) {
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
    var date = new Date();
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
    var currentEditors = spreadsheet.getEditors();
    var currentEditorEmails = [];
    for (var c in currentEditors) {
      currentEditorEmails.push(
        currentEditors[c].getEmail().trim().toLowerCase()
      );
    }

    for (var e in editors) {
      if (currentEditorEmails.indexOf(editors[e]) == -1) {
        spreadsheet.addEditor(editors[e]);
      }
    }
  };
}

//uses helpers.js

function Setting() {
  this.parseDateRange = function (SETTINGS) {
    var YESTERDAY = h.getAdWordsFormattedDate(1, "yyyyMMdd");
    SETTINGS.DATE_RANGE = "20000101," + YESTERDAY;

    if (SETTINGS.DATE_RANGE_LITERAL == "LAST_N_DAYS") {
      SETTINGS.DATE_RANGE =
        h.getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd") + "," + YESTERDAY;
    }

    if (SETTINGS.DATE_RANGE_LITERAL == "LAST_N_MONTHS") {
      var now = new Date(
        Utilities.formatDate(
          new Date(),
          AdsApp.currentAccount().getTimeZone(),
          "MMM dd, yyyy HH:mm:ss"
        )
      );
      now.setHours(12);
      now.setDate(0);

      var TO = Utilities.formatDate(now, "PST", "yyyyMMdd");
      now.setDate(1);

      var counter = 1;
      while (counter < SETTINGS.N) {
        now.setMonth(now.getMonth() - 1);
        counter++;
      }

      var FROM = Utilities.formatDate(now, "PST", "yyyyMMdd");
      SETTINGS.DATE_RANGE = FROM + "," + TO;
    }
  };
}
