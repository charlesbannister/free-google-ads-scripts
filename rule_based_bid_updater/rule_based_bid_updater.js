/*************************************************
* Bid Updater
* Update bids for Search (Shopping & Text Ads)
* Pause/Exclude Keywords/Products
* @version: 1.0.7
* Updates *
 - Labels will be added for raised and lowered bids e.g. "R 9/27". Send an email if there's an error creating a label.
 - 1.0.6 - added Avg. Cpc filter
 - 1.0.7 - added catch for multiple bid updates per run.
 ***************************************************/


let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";

var INPUT_TAB_NAME = "Settings";

var NUMBER_OF_FILTERS = 6;

//if true, only bid keywords/products with a MANUAL_CPC bid strategy will be targeted
//if false, all bid strategies will be targeted (note this will show errors in the logs)
//more info here: https://developers.google.com/adwords/api/docs/appendix/reports/product-partition-report#biddingstrategytype
var EXCLUDE_AUTO_BID_STRATEGIES = true;

var ADMIN_EMAIL = "";

//No need to edit anything below this line

var h = new Helper();
var s = new Setting();

var CUSTOM_METRICS = {
  //metric, metric, operator (divide or multiply, whether high or low is good)
  Ctr: ["Clicks", "Impressions", "divide", "high"],
  Roas: ["conversions_value", "cost", "divide", "high"],
  Cos: ["cost", "conversions_value", "divide", "low"],
  Cpa: ["cost", "Conversions", "divide", "low"],
  AverageCpc: ["cost", "Clicks", "divide", "low"],
  ConversionRate: ["Conversions", "Clicks", "divide", "high"],
  Rpc: ["conversions_value", "Clicks", "divide", "high"]
};

function runScript(SETTINGS, idArray) {
  log("Script Started");
  SETTINGS.CONTROL_SHEET = SpreadsheetApp.openByUrl(
    INPUT_SHEET_URL
  ).getSheetByName(INPUT_TAB_NAME);

  checkSettings(SETTINGS);

  processRowSettings(SETTINGS);

  addLogSheetInfo(SETTINGS);

  log(JSON.stringify(SETTINGS));

  if (SETTINGS.INCLUDE_TEXT) {
    var keywordData = checkKeywords(SETTINGS, idArray);
    var keywordChanges = keywordData["map"];
    var updatedIdArray = keywordData["idArray"];
    updateKeyords(SETTINGS, keywordChanges);
  }

  if (SETTINGS.INCLUDE_SHOPPING) {
    var productGroupData = checkProductGroups(SETTINGS, idArray);
    log(JSON.stringify(productGroupData));
    var productGroupChanges = productGroupData["map"];
    var updatedIdArray = productGroupData["idArray"];
    updateProducts(SETTINGS, productGroupChanges);
  }

  SETTINGS.LOGS.push("The script ran successfully");
  updateControlSheet("", SETTINGS);

  // sendEmail(SETTINGS);

  log("Finished");
  return updatedIdArray;
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
    throw "Error: The label with name '" + str + "' cannot be found";
  }
}

function addLogSheetInfo(SETTINGS) {
  SETTINGS.LOG_SHEET_URL = INPUT_SHEET_URL;

  var logSS = SpreadsheetApp.openByUrl(SETTINGS.LOG_SHEET_URL);
  SETTINGS.logSS = logSS;
  SETTINGS.KEYWORD_LOG_SHEET = logSS.getSheetByName("Keywords");
  SETTINGS.PRODUCT_LOG_SHEET = logSS.getSheetByName("Products");
  // SETTINGS.CHANGE_LOG_SHEET = logSS.getSheetByName("Change Log");
}

function sendEmail(SETTINGS) {
  if (SETTINGS.PREVIEW_MODE) return;

  //Send email
  var SUB =
    AdsApp.currentAccount().getName() + " - " + INPUT_TAB_NAME + " script.";
  var MSG =
    "Hi,<br><br>The " +
    INPUT_TAB_NAME +
    " script ran successfully. Here are the logs:<br><br>";

  MSG += "<ul>";
  for (var l in SETTINGS.LOGS) {
    MSG += "<li>";
    MSG += SETTINGS.LOGS[l];
    MSG += "</li>";
  }
  MSG += "</ul>";

  MSG +=
    "<br><br>Please follow the link below for more information:<br>" +
    SETTINGS.LOG_SHEET_URL;
  MSG += "<br><br>Thanks.";
  var emails = SETTINGS.EMAILS;

  for (var i in emails) {
    MailApp.sendEmail({
      to: emails[i],
      subject: SUB,
      htmlBody: MSG
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

function checkProductGroups(SETTINGS, idArray) {
  let selectStatement =
    "SELECT ad_group_criterion.effective_cpc_bid_micros,ad_group_criterion.display_name, ad_group_criterion.cpc_bid_micros, metrics.clicks, metrics.conversions, metrics.conversions_value, metrics.cost_micros, metrics.impressions, ad_group.name, campaign.name, ad_group.id, campaign.id, ad_group_criterion.listing_group.case_value.product_item_id.value, campaign.bidding_strategy_type, campaign.bidding_strategy, metrics.search_impression_share, metrics.search_absolute_top_impression_share, ad_group_criterion.criterion_id FROM product_group_view";

  var query = [
    selectStatement,
    getQueryWhereString(SETTINGS),
    " and segments.date >" + SETTINGS.DATE_RANGE[0],
    " and segments.date <=" + SETTINGS.DATE_RANGE[1]
  ].join(" ");

  log("product partition report query: " + query);

  var map = { ids: [], rows: {} };
  var reportIter = AdsApp.report(query).rows();
  var number_of_rows = 0;
  while (reportIter.hasNext()) {
    var row = reportIter.next();
    if (
      typeof row["ad_group_criterion.effective_cpc_bid_micros"] ==
      "undefined"
    )
      continue;
    number_of_rows++;
    row["metrics.impressions"] = parseInt(row["metrics.impressions"], 10);
    row["metrics.clicks"] = parseInt(row["metrics.clicks"], 10);
    row["metrics.conversions"] = parseFloat(
      row["metrics.conversions"].toString().replace(/,/g, "")
    );
    row["metrics.cost_micros"] = parseFloat(
      row["metrics.cost_micros"].toString().replace(/,/g, "")
    );
    row["metrics.cost"] = microsToMoney(row["metrics.cost_micros"]);
    row["metrics.conversions_value"] = parseFloat(
      row["metrics.conversions_value"].toString().replace(/,/g, "")
    );
    row["cpc_bid"] = microsToMoney(
      row["ad_group_criterion.cpc_bid_micros"]
    );

    row = addCustomMetricsToRow(row);

    row.newBid = calculateNewBid(row, SETTINGS);

    if (skipEntity(row, SETTINGS)) continue;

    var rowId = row["ad_group.id"] + row["ad_group_criterion.criterion_id"];
    if (idArray.indexOf(rowId) < 0) {
      idArray.push(rowId);
    } else {
      continue;
    }
    map["ids"].push([
      row["ad_group.id"],
      row["ad_group_criterion.criterion_id"]
    ]);
    map["rows"][rowId] = {};
    map["rows"][rowId] = row;
  }
  log(number_of_rows + " initial products returned from the api query");

  //  log(JSON.stringify(map))
  log("Num of product changes: " + map["ids"].length);
  return { map: map, idArray: idArray };
}

// var SETTINGS = {
//     FILTER_1_METRIC:"Clicks", FILTER_1_OPERATOR:">", FILTER_1_VALUE:5,
//     FILTER_2_METRIC:"Conversions", FILTER_2_OPERATOR:">", FILTER_2_VALUE:2,
//     FILTER_3_METRIC:"CPA", FILTER_3_OPERATOR:"<", FILTER_3_VALUE:5,
//     FILTER_4_METRIC:"Labels", FILTER_4_OPERATOR:"IN", FILTER_4_VALUE:"Alpha, Beta",
//     ACTION: "Decrease by amount", CHANGE: .4
// }
// swapLabelTextForIds(SETTINGS)
// log(SETTINGS)
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

// log(getFilterWhereString(SETTINGS))

function getFilterWhereString(SETTINGS) {
  //custom metrics map
  //name : formula, whether high or low is good (e.g. Roas high is good, CPA high is bad - used for Vs target calcs)

  let metricsMap = {
    Cost: "metrics.cost_micros",
    Impressions: "metrics.impressions",
    Clicks: "metrics.clicks",
    Conversions: "metrics.conversions",
    ConversionValue: "metrics.conversion_value",
    CpcBid: "ad_group_criterion.cpc_bid_micros"
  };

  var filters = Object.keys(SETTINGS).map(function (x) {
    if (x.indexOf("FILTER") > -1) {
      return x;
    }
  });

  var filterMap = {};
  var filterParts = ["metric", "operator", "value"];
  var numberOfFilters = filters.length / filterParts.length;

  for (var i = 0; i < numberOfFilters; i++) {
    var filterName = "FILTER_" + (i + 1);
    if (Object.keys(SETTINGS).indexOf(filterName + "_METRIC") == -1)
      continue;
    if (SETTINGS[filterName + "_METRIC"] == "") continue;

    let filterKey = filterName + "_METRIC";
    let sheetMetricName = SETTINGS[filterKey]; //e.g. Clicks, Cost, etc.
    if (Object.keys(CUSTOM_METRICS).indexOf(sheetMetricName) > -1) continue;
    let apiMetricName = metricsMap[sheetMetricName];

    filterMap[filterName] = filterMap[filterName] || {};
    filterMap[filterName]["metric"] = apiMetricName;

    filterKey = filterName + "_OPERATOR";
    filterMap[filterName]["operator"] = SETTINGS[filterKey];

    filterKey = filterName + "_VALUE";
    let value = SETTINGS[filterKey]; //e.g. 50
    if (sheetMetricName == "Cost" || sheetMetricName == "CpcBid") {
      value = value * 1000000; //cost uses micros but we'll allow monetary amounts in the script
    }
    filterMap[filterName]["value"] = value;
  }

  log(JSON.stringify(filterMap));

  var where = "";

  var whereArray = [];

  //turn the filters object into a where statement string
  whereArray.push(filtersToWhereStatement(filterMap, filterParts));
  function filtersToWhereStatement(filterMap, filterParts) {
    var str = [];
    for (var filter in filterMap) {
      //if the metric is a custom metric, continue
      if (
        Object.keys(CUSTOM_METRICS).indexOf(
          filterMap[filter]["metric"]
        ) > -1
      )
        continue;
      if (
        String(filterMap[filter]["value"])
          .toLowerCase()
          .indexOf("avg") > -1
      )
        continue;
      str.push("and");
      for (var p in filterParts) {
        if (filterParts[p] == "metric") {
          str.push(filterMap[filter][filterParts[p]].toLowerCase());
        } else {
          str.push(filterMap[filter][filterParts[p]]);
        }
      }
    }
    return str.join(" ");
  }

  where += whereArray.join(" ");

  return where;
}

function getQueryWhereString(SETTINGS) {
  var where =
    "where campaign.status = ENABLED and ad_group.status = ENABLED ";

  if (SETTINGS.FILTER_1_METRIC === "Clicks")
    where +=
      " and metrics.clicks " +
      SETTINGS.FILTER_1_OPERATOR +
      " " +
      SETTINGS.FILTER_1_VALUE;
  //  where +=  ' and BiddingStrategyType = "MANUAL_CPC" '

  var whereArray = [];

  for (var i in SETTINGS.CAMPAIGN_NAME_CONTAINS) {
    whereArray.push(
      " and campaign.name LIKE '" +
      SETTINGS.CAMPAIGN_NAME_CONTAINS[i].trim() +
      "'"
    );
  }

  for (var i in SETTINGS.CAMPAIGN_NAME_NOT_CONTAINS) {
    whereArray.push(
      " and campaign.name NOT LIKE '" +
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
    let firstMetric = `metrics.${CUSTOM_METRICS[
      metricName
    ][0].toLowerCase()}`;
    let secondMetric = `metrics.${CUSTOM_METRICS[
      metricName
    ][1].toLowerCase()}`;
    if (valueUndefined(row[firstMetric]))
      throw "Error: Can't find the metric " + firstMetric;
    if (valueUndefined(row[secondMetric]))
      throw "Error: Can't find the metric " + secondMetric;

    if (CUSTOM_METRICS[metricName][2] == "divide") {
      row[metricName] =
        row[firstMetric] == 0 || row[secondMetric] == 0
          ? 0
          : round(row[firstMetric] / row[secondMetric], 4);
    }
    if (CUSTOM_METRICS[metricName][2] == "multiply") {
      row[metricName] =
        row[firstMetric] == 0 || row[secondMetric] == 0
          ? 0
          : round(row[firstMetric] * row[secondMetric], 4);
    }
  }
  return row;
}

// function addCustomMetricsToCols(cols, SETTINGS){

//     var filter_map= getFilterMap(SETTINGS)
//     var filter_metrics = Object.keys(filter_map).map(function (x){return filter_map[x]["metric"]})
//     for(f in filter_metrics){
//         if(filter_metrics[f].toLowerCase()=="cpa")continue
//         if(filter_metrics[f].toLowerCase()=="ctr")continue
//         if(filter_metrics[f].toLowerCase()=="acpc")continue
//         if(filter_metrics[f].toLowerCase()=="conversionrate")continue
//         if(filter_metrics[f].toLowerCase()=="Roas")continue

//         if(cols.indexOf(filter_metrics[f])===-1){
//             cols.push(filter_metrics[f])
//         }
//     }
//     return cols
// }
// // cols = cols.concat(filter_metrics)
// log(addCustomMetricsToCols(cols, SETTINGS))

function checkKeywords(SETTINGS, idArray) {
  let selectStatement =
    "SELECT  ad_group_criterion.keyword.text,ad_group_criterion.parental_status.type, ad_group_criterion.keyword.match_type, ad_group_criterion.criterion_id, ad_group_criterion.cpc_bid_micros, metrics.top_impression_percentage, ad_group_criterion.labels, ad_group_criterion.effective_cpc_bid_micros, metrics.conversions_value, metrics.clicks, metrics.conversions, metrics.conversions_value, metrics.cost_micros, metrics.impressions, ad_group.name, campaign.name, ad_group.id, campaign.id, campaign.bidding_strategy_type, campaign.bidding_strategy, metrics.search_impression_share, metrics.search_absolute_top_impression_share FROM keyword_view";

  var query = [
    selectStatement,
    getQueryWhereString(SETTINGS),
    " and ad_group_criterion.status = ENABLED ",
    " and segments.date >" + SETTINGS.DATE_RANGE[0],
    " and segments.date <=" + SETTINGS.DATE_RANGE[1]
  ].join(" ");

  log("Keyword query: " + query);
  // let sheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName("Keyword Report")
  //  AdsApp.report(query).exportToSheet(sheet)

  var map = { ids: [], rows: {} };
  var reportIter = AdsApp.report(query).rows();
  var number_of_rows = 0;
  while (reportIter.hasNext()) {
    var row = reportIter.next();

    //rows 3000006 and 3000000 are AutomaticContent and Content, respectively - more info here: https://groups.google.com/forum/#!topic/adwords-api/qcskfkalb3g
    //AutomaticContent: stats from Display Optimiser. Content: All Display Stats combined
    number_of_rows++;
    row["metrics.impressions"] = parseInt(row["metrics.impressions"], 10);
    row["metrics.clicks"] = parseInt(row["metrics.clicks"], 10);
    row["metrics.conversions"] = parseFloat(
      row["metrics.conversions"].toString().replace(/,/g, "")
    );
    row["metrics.cost_micros"] = parseFloat(
      row["metrics.cost_micros"].toString().replace(/,/g, "")
    );
    row["metrics.cost"] = microsToMoney(row["metrics.cost_micros"]);
    row["metrics.conversions_value"] = parseFloat(
      row["metrics.conversions_value"].toString().replace(/,/g, "")
    );
    row["cpc_bid"] = microsToMoney(
      row["ad_group_criterion.cpc_bid_micros"]
    );

    row = addCustomMetricsToRow(row);

    row.newBid = calculateNewBid(row, SETTINGS);

    //  log(JSON.stringify(row))
    if (skipEntity(row, SETTINGS)) {
      continue;
    }

    var rowId = row["ad_group.id"] + row["ad_group_criterion.criterion_id"];
    if (idArray.indexOf(rowId) < 0) {
      idArray.push(rowId);
    } else {
      continue;
    }
    map["ids"].push([
      row["ad_group.id"],
      row["ad_group_criterion.criterion_id"]
    ]);
    map["rows"][rowId] = {};
    map["rows"][rowId] = row;
  }

  //log(JSON.stringify(map))
  log("Num of keyword changes: " + map["ids"].length);
  //  log(number_of_rows + " initial keywords returned from the api query");

  return { map: map, idArray: idArray };
}

//micros to a monetary amount
function microsToMoney(number) {
  return number / 1000000;
}

function getFilterMap(SETTINGS, row) {
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
    if (Object.keys(SETTINGS).indexOf(filterName + "_METRIC") == -1)
      continue;
    if (SETTINGS[filterName + "_METRIC"] == "") continue;
    filterMap[filterName] = filterMap[filterName] || {};
    for (var x in filterParts) {
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
  //  if (row['ad_group_criterion.cpc_bid_micros'].trim() == "--") {
  //    return true;
  //  }

  if (
    EXCLUDE_AUTO_BID_STRATEGIES &&
    row["campaign.bidding_strategy_type"] !== "MANUAL_CPC" &&
    row["campaign.bidding_strategy_type"] !== "--"
  ) {
    // log("Skipping keyword id " + row.Id)
    return true;
  }

  //custom metrics map
  //name : formula, whether high or low is good (e.g. Roas high is good, CPA high is bad - used for Vs target calcs)
  var filterMap = getFilterMap(SETTINGS, row);

  // log(JSON.stringify(filterMap))

  for (var filter in filterMap) {
    var this_filter = filterMap[filter];
    if (filterNotInCustomMetrics(this_filter.metric)) continue;
    var eval_string =
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
    if (metric == "ad_group_criterion.effective_cpc_bid_micro")
      return false;
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

function calculateNewBid(row, SETTINGS) {
  var currentBid = parseFloat(row["cpc_bid"]);
  var newBid = currentBid;

  //  if(isNaN(currentBid)){
  //   log(JSON.stringify(row))
  //  }
  //  log("Calculating bid. Start bid: " + currentBid)

  var actions = {
    increase_amount: "Increase by amount",
    increase_percent: "Increase by %",
    decrease_amount: "Decrease by amount",
    decrease_percent: "Decrease by %",
    pause: "Pause"
  };

  var action_values = Object.keys(actions).map(function (x) {
    return actions[x];
  });

  if (action_values.indexOf(SETTINGS.ACTION) === -1) {
    log("The action must be one of: " + action_values);
    throw "Error: The action isn't recognised please check the sheet";
  }

  if (SETTINGS.ACTION === actions["increase_amount"]) {
    newBid = currentBid + SETTINGS.CHANGE;
  }

  if (SETTINGS.ACTION === actions["increase_percent"]) {
    newBid = currentBid * (1 + SETTINGS.CHANGE);
  }

  if (SETTINGS.ACTION === actions["decrease_amount"]) {
    newBid = currentBid - SETTINGS.CHANGE;
  }

  if (SETTINGS.ACTION === actions["decrease_percent"]) {
    newBid = currentBid * (1 - SETTINGS.CHANGE);
  }

  //  log("New bid before min/max check: " + newBid)

  if (SETTINGS.MIN_BID !== "" && newBid < SETTINGS.MIN_BID)
    newBid = SETTINGS.MIN_BID;
  if (SETTINGS.MAX_BID !== "" && newBid > SETTINGS.MAX_BID)
    newBid = SETTINGS.MAX_BID;

  //  log("New bid after min/max check: " + newBid)

  return parseFloat(newBid);
}

function updateProducts(SETTINGS, productGroupChanges) {
  if (productGroupChanges["ids"].length == 0) {
    log("No products to update.");
    return;
  }

  var bidLogArray = [];
  var pauseLogArray = [];

  var chunkedArray = [];
  var chunkSize = 10000;

  for (var i = 0; i < productGroupChanges["ids"].length; i += chunkSize) {
    chunkedArray.push(productGroupChanges["ids"].slice(i, i + chunkSize));
  }

  for (var i = 0; i < chunkedArray.length; i++) {
    var productGroups = AdsApp.productGroups()
      .withIds(chunkedArray[i])
      .get();
    while (productGroups.hasNext()) {
      var productGroup = productGroups.next();
      var productGroupId = productGroup.getId();
      var adGroup = productGroup.getAdGroup();
      var rowId = String(adGroup.getId()) + String(productGroupId);
      var row = productGroupChanges["rows"][rowId];

      var oldBid = row.cpc_bid;
      if (oldBid === "Excluded") continue;
      var target = SETTINGS["TARGET_" + SETTINGS.TARGET_TYPE];
      var actual = row[SETTINGS.TARGET_TYPE];
      var actualVsTarget = getActualVsTarget(
        SETTINGS.TARGET_TYPE,
        actual,
        target
      );
      var changePercentage = getChangePercentage(oldBid, row.newBid);
      var newBid = row.newBid;

      //  log("Bid strategy: " + row.BiddingStrategyType)
      if (
        EXCLUDE_AUTO_BID_STRATEGIES &&
        row["campaign.bidding_strategy_type"] !== "MANUAL_CPC"
      ) {
        // log("Skipping product id " + row.Id +". Bid strategy: " + row.BiddingStrategyType)
        continue;
      }

      var preview_mode_text = SETTINGS.PREVIEW_MODE
        ? "(Preview Mode) "
        : "";
      var action =
        SETTINGS.ACTION == "Pause"
          ? preview_mode_text + "Pause"
          : preview_mode_text +
          SETTINGS.ACTION +
          " (" +
          SETTINGS.CHANGE +
          ")";

      var logRow = [
        row["campaign.name"],
        row["ad_group.name"],
        row["ad_group_criterion.display_name"],
        row["metrics.clicks"],
        row["AverageCpc"],
        row["Ctr"],
        row["metrics.impressions"],
        row["metrics.cost"],
        row["metrics.conversions"],
        row["Cpa"],
        row["metrics.conversions_value"],
        row["Roas"],
        row["Cos"],
        row["metrics.search_absolute_top_impression_share"],
        row["cpc_bid"],
        row.newBid,
        action,
        SETTINGS.NOW
      ];

      if (SETTINGS.ACTION === "Pause") {
        productGroup.exclude();
        newBid = SETTINGS.ACTION;
        // log("Excluding "+ row.Id)
      } else {
        // log("Updating the bid of "+ row.Id +" from " + row.CpcBid + " to " + row.newBid)
        try {
          productGroup.setMaxCpc(row.newBid);
        } catch (e) {
          log(
            "Error updating product id " +
            productGroupId +
            " bid to " +
            row.newBid +
            ". Bid strategy: " +
            row.BiddingStrategyType
          );
        }
      }

      if (newBid === "Pause") {
        pauseLogArray.push(logRow);
      } else {
        bidLogArray.push(logRow);
      }
    }
  }

  writeToSheet(SETTINGS, bidLogArray, "Products");
  log(parseInt(bidLogArray.length) + " product bids updated");
  writeToSheet(SETTINGS, pauseLogArray, "Products");
  log(parseInt(pauseLogArray.length) + " products excluded");
}

function updateKeyords(SETTINGS, keywordChanges) {
  log("Updating " + keywordChanges["ids"].length + " keywords...");
  //  log(keywordChanges["ids"])
  //  log(JSON.stringify(keywordChanges["rows"]));

  var bidLogArray = [];
  var pauseLogArray = [];

  var chunkedArray = [];
  var chunkSize = 10000;

  for (var i = 0; i < keywordChanges["ids"].length; i += chunkSize) {
    chunkedArray.push(keywordChanges["ids"].slice(i, i + chunkSize));
  }

  for (var i = 0; i < chunkedArray.length; i++) {
    //  log(chunkedArray[i]);
    var keywords = AdsApp.keywords()
      .withIds(chunkedArray[i])
      .get();
    while (keywords.hasNext()) {
      var keyword = keywords.next();
      var keywordId = keyword.getId();
      var adGroupId = keyword.getAdGroup().getId();
      var rowId = String(adGroupId) + String(keywordId);

      var row = keywordChanges["rows"][rowId];

      var oldBid = row.cpc_bid;
      var target = SETTINGS["TARGET_" + SETTINGS.TARGET_TYPE];
      var actual = row[SETTINGS.TARGET_TYPE];
      var actualVsTarget = getActualVsTarget(
        SETTINGS.TARGET_TYPE,
        actual,
        target
      );
      var changePercentage = getChangePercentage(oldBid, row.newBid);
      var newBid = row.newBid;

      var preview_mode_text = SETTINGS.PREVIEW_MODE
        ? "(Preview Mode) "
        : "";
      var action =
        SETTINGS.ACTION == "Pause"
          ? preview_mode_text + "Pause"
          : preview_mode_text +
          SETTINGS.ACTION +
          " (" +
          SETTINGS.CHANGE +
          ")";

      var logRow = [
        row["campaign.name"],
        row["ad_group.name"],
        row["ad_group_criterion.keyword.text"],
        row["ad_group_criterion.keyword.match_type"],
        row["metrics.clicks"],
        row["AverageCpc"],
        row["Ctr"],
        row["metrics.impressions"],
        row["metrics.cost"],
        row["metrics.conversions"],
        row["Cpa"],
        row["metrics.conversions_value"],
        row["Roas"],
        row["Cos"],
        row["metrics.top_impression_percentage"],
        row["metrics.search_absolute_top_impression_share"],
        row["ad_group_criterion.labels"],
        row["cpc_bid"],
        row.newBid,
        action,
        SETTINGS.NOW
      ];
      //  if (EXCLUDE_AUTO_BID_STRATEGIES && row.BiddingStrategyType !== "cpc") {
      //    // log("Skipping keyword id " + row.Id)
      //    continue;
      //  }

      if (SETTINGS.ACTION === "Pause") {
        keyword.pause();
        newBid = SETTINGS.ACTION;
        changePercentage = "";
      } else {
        try {
          keyword.bidding().setCpc(row.newBid);
        } catch (e) {
          log(
            "Error updating keyword id " +
            keywordId +
            " bid to " +
            row.newBid +
            ". Bid Strategy: " +
            row.BiddingStrategyType
          );
          //  log(JSON.stringify(row))
          continue;
        }

        //  addChangeLabel(keyword, row.CpcBid, row.newBid);
      }

      if (newBid === "Pause") {
        pauseLogArray.push(logRow);
      } else {
        bidLogArray.push(logRow);
      }
    }
  }

  writeToSheet(SETTINGS, bidLogArray, "Keywords");
  log(parseInt(bidLogArray.length) + " keyword bids updated");
  writeToSheet(SETTINGS, pauseLogArray, "Keywords");
  log(parseInt(pauseLogArray.length) + " keywords paused");
}

function addChangeLabel(keyword, old_bid, new_bid) {
  if (old_bid === new_bid) return;

  var suffix;
  if (old_bid > new_bid) suffix = "L";
  if (old_bid < new_bid) suffix = "R";

  var date = new Date();

  var month = String(date.getMonth() + 1);
  var day = String(date.getDate());

  var label = suffix + " " + month + "/" + day;

  checkLabel(label);
  keyword.applyLabel(label);
}

function isToday(date) {
  return date.getDate() == new Date().getDate();
}

function writeArrayToSheet(array, sheet, start_row) {
  sheet
    .getRange(start_row, 1, array.length, array[0].length)
    .setValues(array);
}
function writeToSheet(SETTINGS, logArray, sheetName) {
  if (logArray.length === 0) return;
  // log("Adding "+ logArray.length + " changes to " + sheetName)
  var sheet = SETTINGS.logSS.getSheetByName(sheetName);
  sheet
    .getRange("A2")
    .setValue(
      "Current Data For Lookback (" +
      SETTINGS.N +
      " days)" +
      " - " +
      SETTINGS.DATE_RANGE
    );
  sheet.insertRowsAfter(3, logArray.length);
  var max_rows = 20000;
  if (sheet.getLastRow() > max_rows) {
    sheet.deleteRows(max_rows, sheet.getLastRow() - max_rows);
  }

  writeArrayToSheet(logArray, sheet, 4);

  //   sheet.getRange(1,2,sheet.getLastRow(),sheet.getLastColumn()).setNumberFormat("0.00")
  //   sheet.getRange(1,13,sheet.getLastRow(),sheet.getLastColumn()).setNumberFormat("0.00%")
  //   sheet.getRange(1,7,sheet.getLastRow(),1).setNumberFormat("0.00%")
  //  sheet.getRange(1,10,sheet.getLastRow(),1).setNumberFormat("0.00%")
}

function getChangePercentage(before, after) {
  if (before === 0 && after === 0) return 0;
  return (after - before) / before;
}

function parseDateRange(SETTINGS) {
  var YESTERDAY = getAdWordsFormattedDate(1, "yyyyMMdd");

  SETTINGS.DATE_RANGE = [
    getAdWordsFormattedDate(SETTINGS.N, "yyyyMMdd"),
    YESTERDAY
  ];
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
  var SINGLE_ACCOUNT_HEADER_TYPES = {
    NAME: "normal",
    EMAILS: "csv",
    FLAG: "bool",
    N: "normal",
    ACTION: "normal",
    CHANGE: "normal",
    CAMPAIGN_NAME_CONTAINS: "csv",
    CAMPAIGN_NAME_NOT_CONTAINS: "csv"
  };

  var FILTER_HEADER_TYPES = getFilterHeaderTypes();

  SINGLE_ACCOUNT_HEADER_TYPES2 = {
    MIN_BID: "normal",
    MAX_BID: "normal",
    INCLUDE_SHOPPING: "bool",
    INCLUDE_TEXT: "bool",
    LOG_SHEET_URL: "normal",
    LOGS_COLUMN: "normal"
  };

  var HEADER_TYPES = objectMerge(
    SINGLE_ACCOUNT_HEADER_TYPES,
    FILTER_HEADER_TYPES,
    SINGLE_ACCOUNT_HEADER_TYPES2
  );

  return HEADER_TYPES;
}

function scanForAccounts() {
  log("getting settings...");
  // log(
  //     "The settings sheet should contain " +
  //         NUMBER_OF_FILTERS +
  //         " filter sets"
  // );
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
  //  log(data);

  HEADER_TYPES = getHeaderTypes();

  // log(JSON.stringify(HEADER_TYPES))

  var HEADER = Object.keys(HEADER_TYPES);
  //  log(HEADER);
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
    if (data[k][flagPosition].toLowerCase() != "yes") {
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

function isList(operator) {
  var list_operators = [
    "IN",
    "NOT_IN",
    "CONTAINS_ANY",
    "CONTAINS_NONE",
    "CONTAINS_ALL"
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
      throw "error setting type " + type + " not recognised for " + key;
  }
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
  SETTINGS.PREVIEW_MODE = AdsApp.getExecutionInfo().isPreview();

  if (SETTINGS.PREVIEW_MODE) {
    var msg = "Running in preview mode. No changes will be made.";
    SETTINGS.LOGS.push(msg);
  }
}

/**
 * Checks the settings for issues
 * @returns nothing
 **/
function checkSettings(SETTINGS) {
  //check the settings here

  if (SETTINGS.MAX_BID === "")
    updateControlSheet("Please set a max bid", SETTINGS);
  if (SETTINGS.MIN_BID === "")
    updateControlSheet("Please set a min bid", SETTINGS);
  if (SETTINGS.N === "")
    updateControlSheet("Please set a lookback window", SETTINGS);
}

function checkLabel(labelName) {
  //if the label from the sheet doesn't exist, create it

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
    try {
      var colour = labelName.indexOf("R") > -1 ? "green" : "red";
      AdsApp.createLabel(labelName, "", colour);
    } catch (e) {
      sendLabelError(labelName, e);
    }
  }
}

function sendLabelError(labelName, error) {
  var SUB = "Script Error: Problem Adding Label";

  var MSG = "Hi,<br><br>";
  MSG += "There was a problem adding a label. Here are the details:<br><br>";

  MSG += "<h3>Account</h3>";
  MSG += "<p>" + AdsApp.currentAccount().getName() + "</p>";
  MSG += "<h3>Script</h3>";
  MSG += "<p>" + INPUT_TAB_NAME + "</p>";
  MSG += "<h3>Label Text</h3>";
  MSG += "<p>" + labelName + "</p>";
  MSG += "<h3>Error Message (from Google Ads)</h3>";
  MSG += "<p>" + error + "</p>";

  MSG += "<br><br>The settings sheet is here:<br>" + INPUT_SHEET_URL;
  MSG += "<br><br>Thanks.";

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: SUB,
    htmlBody: MSG
  });
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
    currentEditorEmails.push(
      currentEditors[c]
        .getEmail()
        .trim()
        .toLowerCase()
    );
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
    LOGS_COLUMN =
      controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) {
      break;
    }
    col++;
  }
  return LOGS_COLUMN;
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
  console.log(AdsApp.currentAccount().getName() + " - " + msg);
}

function round(num, n) {
  return +(Math.round(parseFloat(num) + "e+" + n) + "e-" + n);
}

function runAccount() {
  log("Account running");
}

function runRows(INPUT) {
  log("running rows");
  var SETTINGS = JSON.parse(INPUT)[
    AdsApp.currentAccount()
      .getCustomerId()
      .toString()
  ];
  var idArray = [];
  for (var rowId in SETTINGS) {
    var idArray = runScript(SETTINGS[rowId], idArray);
    //  log(idArray.length);
  }
}

function callBack() {
  // Do something here
  Logger.log("Finished");
}

function main() {
  if (isMCC()) {
    var SETTINGS = scanForAccounts();
    log(JSON.stringify(SETTINGS));
    var ids = Object.keys(SETTINGS);
    //  ids = ids.map((id)=>{return id.split("-").join("")})
    log(`Account ids to run: ${ids}`);
    if (ids.length == 0) {
      Logger.log("No Rules Specified");
      return;
    }
    AdsManagerApp.accounts()
      .withIds(ids)
      .withLimit(50)
      .executeInParallel("runRows", "callBack", JSON.stringify(SETTINGS));
  } else {
    var ALL_SETTINGS = scanForAccounts();

    //run all rows and all accounts
    var idArray = [];
    for (var S in ALL_SETTINGS) {
      for (var R in ALL_SETTINGS[S]) {
        var idArray = runScript(ALL_SETTINGS[S][R], idArray);
        log(idArray.length);
      }
    }
  }
}

function isMCC() {
  try {
    AdsManagerApp.accounts();
    return true;
  } catch (e) {
    if (String(e).indexOf("not defined") > -1) {
      return false;
    } else {
      return true;
    }
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