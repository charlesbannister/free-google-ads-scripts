/**
 * Account Search Term NGrams
 * @author Charles Bannister
 * @version 1.0.0
 * Free updates & support at  https://shabba.io/script/4
 * 
**/

// Template: https://docs.google.com/spreadsheets/d/1L6ty0u7OtD3Ed5h4SBN4mCwkbAqYsaw_18kOLQDr6Tk
// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1L6ty0u7OtD3Ed5h4SBN4mCwkbAqYsaw_18kOLQDr6Tk/copy
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";





var INPUT_TAB_NAME = "Settings";

var h = new Helper();
var s = new Setting();

var masterListSheetName = "Already Added";

const SCRIPT_NAME = "nGrams";
const SHABBA_SCRIPT_ID = 4;

function runScript(SETTINGS) {
  processRowSettings(SETTINGS);

  log(JSON.stringify(SETTINGS));

  checkSettings(SETTINGS);

  new setupOutputSheet(SETTINGS);

  addNegatives(SETTINGS);

  let nGramsFound = getQueriesAndWriteToSheet(SETTINGS);

  SETTINGS.LOGS.push("The script ran successfully");
  updateControlSheet("", SETTINGS);

  if (nGramsFound && SETTINGS['NOTIFY']) {
    sendEmail(SETTINGS);
  }

  log("Finished");
}

function sendEmail(SETTINGS) {
  if (SETTINGS.PREVIEW_MODE) return;
  log("Sending email")

  //Send email
  var SUB =
    SETTINGS['NAME'] +
    " - nGrams Script has Results";
  var MSG =
    `Hi,<br><br>nGrams were found for the "${SETTINGS['NAME']}" account
        <br>
        You can view the nGrams here: ${SETTINGS['LOG_SHEET_URL']}
        <br>
        You can view the settings sheet here: ${INPUT_SHEET_URL}
        `;

  var emails = SETTINGS.EMAILS;

  for (var i in emails) {
    MailApp.sendEmail({
      to: emails[i],
      subject: SUB,
      htmlBody: MSG
    });
  }
}

function addNegatives(SETTINGS) {
  log(
    `Adding negative keywords to the '${SETTINGS["NEGATIVE_KEYWORD_LIST"]}' list...`
  );
  var nGrams = getNegativesFromSheet(SETTINGS);

  log(nGrams);

  if (nGrams.length == 0) return;

  let negativeList = getNegativeList(SETTINGS["NEGATIVE_KEYWORD_LIST"]);

  let negativeKeywords = [];

  for (var nGramIndex in nGrams) {
    var row = nGrams[nGramIndex];
    // log(JSON.stringify(row))
    let negativeKeyword = addMatchType(row.nGram, row.matchType);
    negativeKeywords.push(negativeKeyword);
  }
  negativeList.addNegativeKeywords(negativeKeywords);
}

function getNegativeList(listName) {
  var listIter = AdWordsApp.negativeKeywordLists()
    .withCondition("Name = '" + listName + "'")
    .get();

  if (listIter.hasNext()) {
    return listIter.next();
  } else {
    throw "The shared negative list ('" + listName + "') can't be found";
  }
}

function getNegativesFromSheet(SETTINGS) {
  var sheetNames = SETTINGS.tabNames.slice(1);

  var nGrams = [];

  for (var s in sheetNames) {
    var sheetName = sheetNames[s];

    var sheet = SETTINGS.logSS.getSheetByName(sheetName);

    var data = sheet.getDataRange().getValues();
    var header = data.shift();
    let addAsNegativeIndex = header.indexOf("Add as negative?");
    let matchTypeIndex = header.indexOf("Match type");
    // log('addAsNegativeIndex: ' + addAsNegativeIndex)
    for (var d in data) {
      var row = data[d];
      var nGram = String(row[header.indexOf("nGram")]);
      var addBool = row[addAsNegativeIndex];
      var matchType = row[matchTypeIndex];

      if (!addBool) continue;
      nGrams.push({
        nGram,
        matchType
      });
    }
  }

  if (!SETTINGS.PREVIEW_MODE) {
    addCheckedQueriesToMasterSheet(SETTINGS, nGrams);
  } else {
    SETTINGS.LOGS.push(
      "Running in preview mode, so the 'Already Added' list won't be updated"
    );
  }

  return nGrams;
}

function addMatchType(keyword, matchType) {
  matchType = matchType.toLowerCase();
  if (matchType === "exact") return "[" + keyword + "]";
  if (matchType === "phrase") return '"' + keyword + '"';
  if (matchType === "broad") return keyword;

  throw "Match type '" +
  matchType +
  "' not recognised. Please check the settings.";
}

function addCheckedQueriesToMasterSheet(SETTINGS, nGrams) {
  var len = Object.keys(nGrams).length;
  if (len === 0) return;

  SETTINGS.LOGS.push(
    "Adding " +
    String(parseInt(len)) +
    " checked nGrams to the 'Already Added' list"
  );

  var logArray = [];

  for (var nGramIndex in nGrams) {
    var row = nGrams[nGramIndex];
    let negativeKeyword = addMatchType(row.nGram, row.matchType);
    logArray.push([negativeKeyword, SETTINGS["NEGATIVE_KEYWORD_LIST"]]);
  }

  var sheet = SETTINGS.logSS.getSheetByName(masterListSheetName);
  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      logArray.length,
      logArray[0].length
    )
    .setValues(logArray);
}

function getQueryWhereString(SETTINGS) {
  var where = "where CampaignStatus in [ENABLED,PAUSED,REMOVED]  ";

  var whereArray = [];

  for (var i in SETTINGS.AD_GROUP_NAME_CONTAINS) {
    whereArray.push(
      " and AdGroupName CONTAINS_IGNORE_CASE '" +
      SETTINGS.AD_GROUP_NAME_CONTAINS[i].trim() +
      "'"
    );
  }

  if (SETTINGS.CAMPAIGN_NAME_EQUALS === "") {
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
  } else {
    where +=
      " and CampaignName = '" +
      SETTINGS.CAMPAIGN_NAME_EQUALS.trim() +
      "'";
  }

  for (var i in SETTINGS.QUERY_NOT_CONTAINS) {
    whereArray.push(
      " and Query DOES_NOT_CONTAIN_IGNORE_CASE '" +
      SETTINGS.QUERY_NOT_CONTAINS[i].trim() +
      "'"
    );
  }

  if (String(SETTINGS["MIN_QUERY_CLICKS"]).trim() != "") {
    whereArray.push(`and Clicks > ${SETTINGS["MIN_QUERY_CLICKS"]}`);
  }

  if (String(SETTINGS["MIN_QUERY_IMPRESSIONS"]).trim() != "") {
    whereArray.push(
      `and Impressions > ${SETTINGS["MIN_QUERY_IMPRESSIONS"]}`
    );
  }

  where += whereArray.join(" ");

  return where;
}

function writePotentialQueriesToSheet(SETTINGS, potentialQueries, sheet) {
  sheet.clear();

  var checkBoxCol = 3;

  sheet
    .getRange(2, checkBoxCol, sheet.getMaxRows(), 1)
    .clearDataValidations();

  if (potentialQueries.length < 2) return;

  //write the data

  sheet
    .getRange(1, 1, potentialQueries.length, potentialQueries[0].length)
    .setValues(potentialQueries);

  //sort by cost
  sheet.sort(8, false);

  //format
  sheet
    .getRange(1, 5, sheet.getLastRow(), 6)
    .setNumberFormat("0");
  sheet
    .getRange(1, 8, sheet.getLastRow(), sheet.getLastColumn())
    .setNumberFormat("0.00");
  sheet.getRange(2, 7, sheet.getLastRow(), 1).setNumberFormat("0.00%");

  var enforceCheckbox = SpreadsheetApp.newDataValidation();
  enforceCheckbox.requireCheckbox();
  enforceCheckbox.setAllowInvalid(true);
  enforceCheckbox.build();

  var range = sheet.getRange(2, checkBoxCol, potentialQueries.length - 1, 1);
  range.setDataValidation(enforceCheckbox);

  var enforceDropdown = SpreadsheetApp.newDataValidation();
  enforceDropdown.requireValueInList(["Exact", "Phrase", "Broad"], true);
  enforceDropdown.setAllowInvalid(true);
  enforceDropdown.build();
  range = sheet.getRange(2, checkBoxCol + 1, potentialQueries.length - 1, 1);
  range.setDataValidation(enforceDropdown);
}

/**
 * Returns true if ngrams were found
 */
function getQueriesAndWriteToSheet(SETTINGS) {

  var sheetNames = SETTINGS.tabNames;
  sheetNames.shift();

  var alreadyAddedNegatives = SETTINGS.logSS
    .getSheetByName(masterListSheetName)
    .getDataRange()
    .getValues()
    .map(function (x) {
      return String(x[0])
        .replace("[", "")
        .replace("]", "")
        .replace('"', "")
        .replace('"', "")
        .toLowerCase();
    });
  alreadyAddedNegatives.shift();
  alreadyAddedNegatives.shift();
  log("alreadyAddedNegatives: " + alreadyAddedNegatives);

  var map = getNGramMap(SETTINGS);

  let nGramsFound = false

  for (var s in sheetNames) {
    var logArray = [
      [
        "Date",
        "nGram",
        "Add as negative?",
        "Match type",
        "Clicks",
        "Impressions",
        "Ctr",
        "Cost",
        "Conversions",
        "Cpa",
        "Conversion Value",
        "ROAS"
      ]
    ];

    var sheetName = sheetNames[s];
    var sheet = SETTINGS.logSS.getSheetByName(sheetName);
    var nGrams = map[sheetName];
    for (var nGram in nGrams) {
      var row = nGrams[nGram];

      if (
        alreadyAddedNegatives.indexOf(nGram) > -1 &&
        SETTINGS["DONT_LOG_ADDED_NEGATIVES"]
      )
        continue;

      row.CTR =
        row.Impressions == 0
          ? 0
          : round(row.Clicks / row.Impressions, 4);
      row.CPA =
        row.Conversions == 0 ? 0 : round(row.Cost / row.Conversions, 2);
      row.ROAS =
        row.ConversionValue == 0
          ? 0
          : round(row.ConversionValue / row.Cost, 2);
      row.CVR = row.Conversions > 0 ? row.Conversions / row.Clicks : 0;

      if (filterNGram(row, SETTINGS)) continue;
      nGramsFound = true
      var logRow = [
        SETTINGS.NOW,
        String(nGram),
        "",
        SETTINGS.NEGATIVE_KEYWORD_DEFAULT_MATCH_TYPE,
        row.Clicks,
        row.Impressions,
        row.CTR,
        row.Cost,
        row.Conversions,
        row.CPA,
        row.ConversionValue,
        row.ROAS
      ];
      logArray.push(logRow);
    }

    log("Found " + String(parseInt(logArray.length) - 1) + " " + sheetName);
    writePotentialQueriesToSheet(SETTINGS, logArray, sheet);
  }
  return nGramsFound
}

function filterNGram(row, SETTINGS) {
  var filter = false; //whether to filter out the ngram

  if (SETTINGS.MIN_CPA === "" || row.CPA > SETTINGS.MIN_CPA) {
    //leave false
  } else {
    return true;
  }
  if (SETTINGS.MAX_ROAS === "" || row.ROAS <= SETTINGS.MAX_ROAS) {
    //leave false
  } else {
    return true;
  }

  if (
    SETTINGS.MAX_CONVERSIONS === "" ||
    row.Conversions < SETTINGS.MAX_CONVERSIONS
  ) {
    //leave false
  } else {
    return true;
  }

  if (SETTINGS.MIN_CLICKS === "" || row.Clicks >= SETTINGS.MIN_CLICKS) {
    //leave false
  } else {
    return true;
  }

  return filter;
}

function getNGramMap(SETTINGS) {
  var sheetNames = SETTINGS.tabNames;

  var checkedQueries = SETTINGS.logSS
    .getSheetByName(sheetNames[0])
    .getDataRange()
    .getValues()
    .map(function (x) {
      return x[0];
    });

  checkedQueries.shift();

  var OPTIONS = { includeZeroImpressions: false };
  var cols = [
    "Query",
    "ConversionValue",
    "Impressions",
    "Clicks",
    "Cost",
    "Conversions"
  ];

  var reportName = "SEARCH_QUERY_PERFORMANCE_REPORT";

  var query = [
    "select",
    cols.join(","),
    "from",
    reportName,
    getQueryWhereString(SETTINGS),
    "during",
    SETTINGS.DATE_RANGE
  ].join(" ");

  log("Query: " + query);

  var map = { "1-grams": {}, "2-grams": {}, "3-grams": {}, "4-grams": {} };

  let queryCount = 0

  var reportIter = AdWordsApp.report(query, OPTIONS).rows();
  while (reportIter.hasNext()) {
    var row = reportIter.next();
    queryCount++
    row.Impressions = parseInt(row.Impressions, 10);
    row.Clicks = parseInt(row.Clicks, 10);
    row.Conversions = parseFloat(row.Conversions);
    row.Cost = parseFloat(row.Cost.toString().replace(/,/g, ""));
    row.ConversionValue = parseFloat(
      row.ConversionValue.toString().replace(/,/g, "")
    );

    var nGrams = row.Query.split(" ");
    var metrics = [
      "Cost",
      "Impressions",
      "Clicks",
      "Conversions",
      "ConversionValue"
    ];

    for (n in nGrams) {
      //1 grams
      var nGram = nGrams[n];

      if (checkedQueries.indexOf(nGram) > -1) continue;
      map["1-grams"][nGram] = map["1-grams"][nGram] || {};
      for (var m in metrics) {
        map["1-grams"][nGram][metrics[m]] =
          map["1-grams"][nGram][metrics[m]] + row[metrics[m]] ||
          row[metrics[m]];
      }

      //2 words
      if (nGrams[parseInt(n) + 1]) {
        var biGram = nGrams[n] + " " + nGrams[parseInt(n) + 1];
        if (checkedQueries.indexOf(biGram) > -1) continue;
        map["2-grams"] = map["2-grams"] || {};
        map["2-grams"][biGram] = map["2-grams"][biGram] || {};
        for (var m in metrics) {
          map["2-grams"][biGram][metrics[m]] =
            map["2-grams"][biGram][metrics[m]] + row[metrics[m]] ||
            row[metrics[m]];
        }
      }

      //3 words
      if (nGrams[parseInt(n) + 1] && nGrams[parseInt(n) + 2]) {
        var biGram =
          nGrams[n] +
          " " +
          nGrams[parseInt(n) + 1] +
          " " +
          nGrams[parseInt(n) + 2];
        if (checkedQueries.indexOf(biGram) > -1) continue;
        map["3-grams"] = map["3-grams"] || {};
        map["3-grams"][biGram] = map["3-grams"][biGram] || {};
        for (var m in metrics) {
          map["3-grams"][biGram][metrics[m]] =
            map["3-grams"][biGram][metrics[m]] + row[metrics[m]] ||
            row[metrics[m]];
        }
      }

      //4 words
      if (
        nGrams[parseInt(n) + 1] &&
        nGrams[parseInt(n) + 2] &&
        nGrams[parseInt(n) + 3]
      ) {
        var biGram =
          nGrams[n] +
          " " +
          nGrams[parseInt(n) + 1] +
          " " +
          nGrams[parseInt(n) + 2] +
          " " +
          nGrams[parseInt(n) + 3];
        if (checkedQueries.indexOf(biGram) > -1) continue;
        map["4-grams"] = map["4-grams"] || {};
        map["4-grams"][biGram] = map["4-grams"][biGram] || {};
        for (var m in metrics) {
          map["4-grams"][biGram][metrics[m]] =
            map["4-grams"][biGram][metrics[m]] + row[metrics[m]] ||
            row[metrics[m]];
        }
      }
    }
  }

  //log(JSON.stringify(map))
  log(`${queryCount} queries were found for consideration under the date range and settings provided.`)
  return map;
}

function setupOutputSheet(SETTINGS) {
  this.setupTabs = function (logSS, tabName, newSheet) {
    var outputTab = logSS.getSheetByName(tabName);

    if (tabName == masterListSheetName) {
      outputTab
        .getRange("A1")
        .setValue(
          "Previously added negative keywords will be stored here. They can optionally be used to prevent the same nGrams appearing again."
        );
      outputTab.getRange("A2").setValue("Negative Keyword");
      outputTab.getRange("B2").setValue("Negative Keyword List");
    }

    outputTab
      .getRange(1, 1, 1, outputTab.getMaxColumns())
      .setFontWeight("bold");

    var maxColumns = outputTab.getMaxColumns(); //total number of cols

    var lastColumn =
      outputTab.getLastColumn() < 7 ? 7 : outputTab.getLastColumn(); //number of populated cols

    var numCols = maxColumns - lastColumn;

    if (numCols > 0) {
      outputTab.deleteColumns(lastColumn + 1, numCols);
    }
  };

  this.createTabs = function (tabNames, logSS) {
    //attempt to rename
    var logSheets = logSS.getSheets();
    for (var l in logSheets) {
      var logSheet = logSheets[l];
      try {
        logSheet.setName(tabNames[l]);
      } catch (e) { }
    }
    //attempt to create
    for (var t in tabNames) {
      var tabName = tabNames[t];

      try {
        logSS.insertSheet(tabName);
      } catch (e) { }
    }
  };

  var reportName = SETTINGS.NAME + " - nGrams Output Sheet";

  var newSheet = SETTINGS.LOG_SHEET_URL == "";

  if (newSheet) {
    SETTINGS.LOGS.push("Creating new output sheet");

    var ss = SpreadsheetApp.create(reportName);
    SETTINGS.LOG_SHEET_URL = ss.getUrl();
  }

  var logSS = SpreadsheetApp.openByUrl(SETTINGS.LOG_SHEET_URL);

  if (logSS.getName() != reportName) {
    logSS.rename(reportName);
  }

  var editors = getEditorsFromSheet(
    SETTINGS.CONTROL_SHEET,
    SETTINGS.LOGS_COLUMN
  );
  h.addEditors(logSS, editors);

  SETTINGS.tabNames = [
    masterListSheetName,
    "1-grams",
    "2-grams",
    "3-grams",
    "4-grams"
  ];

  this.createTabs(SETTINGS.tabNames, logSS);

  //add days and hours to tabs
  for (var t in SETTINGS.tabNames) {
    this.setupTabs(logSS, SETTINGS.tabNames[t], newSheet);
  }

  SETTINGS.logSS = logSS;

  updateControlSheet("", SETTINGS);
}

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

/*

SETTINGS SECTION

*/

function getHeaderTypes() {
  var MCC_HEADER_TYPES = {};


  MCC_HEADER_TYPES = { ID: "normal" };

  var SINGLE_ACCOUNT_HEADER_TYPES = {
    NAME: "normal",
    EMAILS: "csv",
    NOTIFY: "bool",//notify if there are results?
    FLAG: "bool",
    NEGATIVE_KEYWORD_LIST: "normal",
    NEGATIVE_KEYWORD_DEFAULT_MATCH_TYPE: "normal",
    DONT_LOG_ADDED_NEGATIVES: "normal",
    NEGATIVE_KEYWORD_LIST: "normal",
    N: "normal",
    MIN_QUERY_IMPRESSIONS: "normal",
    MIN_QUERY_CLICKS: "normal",
    CAMPAIGN_NAME_CONTAINS: "csv",
    CAMPAIGN_NAME_NOT_CONTAINS: "csv",
    CAMPAIGN_NAME_EQUALS: "normal",
    QUERY_NOT_CONTAINS: "csv"
  };

  SINGLE_ACCOUNT_HEADER_TYPES2 = {
    MIN_CLICKS: "normal",
    MAX_CONVERSIONS: "normal",
    MIN_CPA: "normal",
    MAX_ROAS: "normal",
    PLACEHOLDER_1: "normal",
    PLACEHOLDER_2: "normal",
    PLACEHOLDER_3: "normal",
    PLACEHOLDER_4: "normal",
    PLACEHOLDER_5: "normal",
    LOG_SHEET_URL: "normal",
    LOGS_COLUMN: "normal"
  };

  var HEADER_TYPES = objectMerge(
    MCC_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES2
  );

  return HEADER_TYPES;
}

function scanForAccounts() {
  log("getting settings...");

  var map = {};
  var controlSheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(
    INPUT_TAB_NAME
  );
  var data = SpreadsheetApp.openByUrl(INPUT_SHEET_URL)
    .getSheetByName(INPUT_TAB_NAME)
    .getDataRange()
    .getValues();
  data.shift();
  data.shift();
  data.shift();
  //log(data)

  HEADER_TYPES = getHeaderTypes();

  // log(JSON.stringify(HEADER_TYPES))

  var HEADER = Object.keys(HEADER_TYPES);

  var LOGS_COLUMN = 0;
  var col = 5;
  while (controlSheet.getRange(3, col).getValue()) {
    LOGS_COLUMN =
      controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) {
      break;
    }
    col++;
  }

  var flagPosition = HEADER.indexOf("FLAG");

  for (var k in data) {
    //if "run script" is not set to "yes", continue.

    if (!data[k][flagPosition]) {
      continue;
    }
    var rowNum = parseInt(k, 10) + 4;
    var id = data[k][0];
    var rowId = id + "/" + rowNum;
    map[id] = map[id] || {};
    map[id][rowId] = { ROW_NUM: parseInt(k, 10) + 4 };
    for (var j in HEADER) {
      if (HEADER[j] == "LOGS_COLUMN") {
        map[id][rowId][HEADER[j]] = LOGS_COLUMN;
        continue;
      }
      map[id][rowId][HEADER[j]] = data[k][j];
    }
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

function processSetting(key, value, HEADER, CONTROL_SHEET) {
  var type = HEADER[key];
  if (key == "ROW_NUM") {
    return value;
  }
  switch (type) {
    case "label":
      return [
        CONTROL_SHEET.getRange(
          3,
          Object.keys(HEADER).indexOf(key) + 1
        ).getValue(),
        value
      ];
      break;
    case "normal":
      return value;
      break;
    case "bool"://checkbox
      return value;
      break;
    case "number":
      return value == "" ? 0 : value;
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
    AdWordsApp.currentAccount().getTimeZone(),
    "MMM dd, yyyy HH:mm:ss"
  );

  SETTINGS.CONTROL_SHEET = SpreadsheetApp.openByUrl(
    INPUT_SHEET_URL
  ).getSheetByName(INPUT_TAB_NAME);
  SETTINGS.LOGS_COLUMN = h.getLogsColumn(SETTINGS.CONTROL_SHEET);
  SETTINGS.LOGS = [];

  // log(JSON.stringify(SETTINGS))

  var defaultNote =
    "Possible problems include: 1) There was an error (check the logs within Google Ads) 2) The script was stopped before completion";
  SETTINGS.CONTROL_SHEET.getRange(
    SETTINGS.ROW_NUM,
    SETTINGS.LOGS_COLUMN,
    1,
    1
  )
    .setValue(
      "The script is either still running or didn't finish successfully"
    )
    .setNote(defaultNote);

  parseDateRange(SETTINGS);
  SETTINGS.PREVIEW_MODE = AdWordsApp.getExecutionInfo().isPreview();
}

function parseDateRange(SETTINGS) {
  var YESTERDAY = getAdWordsFormattedDate(1, "yyyyMMdd");

  SETTINGS.DATE_RANGE =
    getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd") + "," + YESTERDAY;
}

/**
 * Checks the settings for issues
 * @returns nothing
 **/
function checkSettings(SETTINGS) {
  //check the settings here
}

function log(msg) {
  Logger.log(AdWordsApp.currentAccount().getName() + " - " + msg);
}

function round(num, n) {
  return +(Math.round(num + "e+" + n) + "e-" + n);
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
  var put = [SETTINGS.LOG_SHEET_URL, logString, SETTINGS.NOW];
  SETTINGS.CONTROL_SHEET.getRange(
    SETTINGS.ROW_NUM,
    SETTINGS.LOGS_COLUMN - 1,
    1,
    3
  ).setValues([put]);
  SETTINGS.CONTROL_SHEET.getRange(
    SETTINGS.ROW_NUM,
    SETTINGS.LOGS_COLUMN,
    1,
    1
  ).setNote(logString);
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
    AdWordsApp.currentAccount().getTimeZone(),
    format
  );
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
    var currentEditors = spreadsheet.getEditors();
    var currentEditorEmails = [];
    for (var c in currentEditors) {
      currentEditorEmails.push(
        currentEditors[c]
          .getEmail()
          .trim()
          .toLowerCase()
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
  this.processSetting = function (key, value, HEADER, controlSheet) {
    var type = HEADER[key];
    if (key == "ROW_NUM") {
      return value;
    }
    var h = new Helper();
    switch (type) {
      case "number":
        if (h.isNumber(value)) {
          return value;
        } else {
          throw "Error: Expected a number but recieved '" +
          value +
          "' for the key '" +
          key +
          "'. Please check the settings";
        }
        return value;
        break;
      case "label":
        return [
          controlSheet
            .getRange(3, Object.keys(HEADER).indexOf(key) + 1)
            .getValue(),
          value
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
        throw "error setting type " +
        type +
        " not recognised for " +
        key;
    }
  };

  this.parseDateRange = function (SETTINGS) {
    var YESTERDAY = h.getAdWordsFormattedDate(1, "yyyyMMdd");
    SETTINGS.DATE_RANGE = "20000101," + YESTERDAY;

    if (SETTINGS.DATE_RANGE_LITERAL == "LAST_N_DAYS") {
      SETTINGS.DATE_RANGE =
        h.getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd") +
        "," +
        YESTERDAY;
    }

    if (SETTINGS.DATE_RANGE_LITERAL == "LAST_N_MONTHS") {
      var now = new Date(
        Utilities.formatDate(
          new Date(),
          AdWordsApp.currentAccount().getTimeZone(),
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

/**
 * It's easy to copy and paste a row and duplicate a Url
 * this means the same Url will be overwritten
 * Add a warning to the logs if this happens
 */
function checkForDuplicateOutputSpreadsheetUrls() {
  const controlSheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME);
  const urls = controlSheet?.getRange(4, 25, controlSheet.getLastRow(), 1)
    .getValues()
    .flat()
    .filter(x => x !== '');
  let findDuplicates = array => array.filter((item, index) => array.indexOf(item) !== index)
  const duplicateUrls = [...new Set(findDuplicates(urls))]
  if (duplicateUrls.length > 0) {
    throw new Error(`Duplicate Log Sheet urls found. Log Sheet Urls must be unique or they will be overwritten. Please check the following urls: ${duplicateUrls.join(', ')}`);
  }
}

function main() {
  try {
    runTopLevelLogic();
  } catch (error) {
    sendErrorEmailToAdmin(error.stack);
    throw new Error(error.stack);
  }
}

function sendErrorEmailToAdmin(error) {
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
  // MailApp.sendEmail(adminEmail, subject, body);
}

function runTopLevelLogic() {
  checkForDuplicateOutputSpreadsheetUrls();
  if (isMCC()) {
    var SETTINGS = scanForAccounts();
    //   log(JSON.stringify(SETTINGS))
    var ids = Object.keys(SETTINGS);
    if (ids.length == 0) {
      Logger.log("No Rules Specified");
      return;
    }
    MccApp.accounts()
      .withIds(ids)
      .withLimit(50)
      .executeInParallel("runRows", "callBack", JSON.stringify(SETTINGS));
    return;
  }
  var ALL_SETTINGS = scanForAccounts();

  //run all rows and all accounts
  for (var S in ALL_SETTINGS) {
    for (var R in ALL_SETTINGS[S]) {
      runScript(ALL_SETTINGS[S][R]);
    }
  }
}

function isMCC() {
  try {
    MccApp.accounts();
    return true;
  } catch (e) {
    if (String(e).indexOf('not defined') > -1) {
      return false;
    } else {
      return true;
    }
  }
}

function runRows(INPUT) {
  log("running rows");
  var SETTINGS = JSON.parse(INPUT)[
    AdWordsApp.currentAccount()
      .getCustomerId()
      .toString()
  ];
  for (var rowId in SETTINGS) {
    runScript(SETTINGS[rowId]);
  }
}

function callBack() {
  // Do something here
  Logger.log("Finished");
}
