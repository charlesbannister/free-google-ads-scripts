/**
 * Meet our Advanced Anomaly Detector for Google Ads
 * @author: Charles Bannister of shabba.io
 * @version 1.2.0
 * Updates:
 *  - 1.1.0: Added weekend notifications option
 *  - 1.2.0: Made it possible to run without alert settings (leave the Alert Parameters operator blank)
 * Purchased from shabba.io and not for resale or redistribution - thanks!
 * Free updates & support at  https://shabba.io/script/13
**/

// Template: https://docs.google.com/spreadsheets/d/1Fcaq3PGgBpSwTqs-PAWEo2lkP7WgyQ4NTRZjC4yuIhI
// Template: https://docs.google.com/spreadsheets/d/1Fcaq3PGgBpSwTqs-PAWEo2lkP7WgyQ4NTRZjC4yuIhI
// File > Make a copy or visit https://docs.google.com/spreadsheets/d/1Fcaq3PGgBpSwTqs-PAWEo2lkP7WgyQ4NTRZjC4yuIhI/copy
let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";


// Getting a Slack Webhook URL is dead easy
// How-to with screenshots here: https://docs.google.com/document/d/1g1CX6ZRMtmx6KjNTjxMutqZp5vca7ESs-I0ztl0LcEM/edit
const SLACK_TEAM_WEBHOOK_URL = '';


// Google Drive Folder ID
// Add an ID for the folder where you want to store the reports
// If unspecified, the root folder will be used
const FOLDER_ID = '';
let INPUT_TAB_NAME = 'Anomaly Detector';

const SCRIPT_NAME = 'Anomaly Detector';
const SHABBA_SCRIPT_ID = 13;

//If true, the settings and output sheets will be set to "anyone can edit"
const ANYONE_CAN_EDIT_SHEETS = false;


let DEBUG = false;

const SETTINGS_TEMPLATE_URL = 'https://docs.google.com/spreadsheets/d/1EXuhMGTGvtRsq2KBKuPx4aNt_bBp-Q-GRobGaBwGC0Y/edit#gid=828446083';

const APP_ENVIRONMENTS = { 'PRODUCTION': 'PRODUCTION' }
const APP_ENVIRONMENT = APP_ENVIRONMENTS.PRODUCTION;


let SENT_SLACK_NOTIFICATIONS = [];


const ALERT_OPERATOR_ENUMS = {
  'INCREASE': 'Increases by',
  'DECREASE': 'Decreases by',
}

const settingsColumnNumbers = {
  'attributeColumn': 1,
  'attributeOperatorColumn': 2,
  'attributeParamColumn': 3,
  'thresholdMetricColumn_prev': 5,
  'thresholdOperatorColumn_prev': 6,
  'thresholdParamColumn_prev': 7,
  'thresholdMetricColumn_curr': 9,
  'thresholdOperatorColumn_curr': 10,
  'thresholdParamColumn_curr': 11,
  'alertMetricColumn': 13,
  'alertOperatorColumn': 14,
  'alertParamColumn': 15,
}

const reportTypes = {
  'AdGroup': 'AdGroup',
  'Campaign': 'Campaign',
  'Label': 'Label',
  'Account': 'Account',
}

let EMAIL_SENT_TIME_COLUMN_NUMBER;
let SLACK_SENT_TIME_COLUMN_NUMBER;

const SLACK_ADMIN_WEBHOOK_URL = '';


Date.prototype.yyyymmdd = function () {
  let yyyy = this.getFullYear().toString();
  let mm = (this.getMonth() + 1).toString(); // getMonth() is zero-based
  let dd = this.getDate().toString();
  return yyyy + (mm[1] ? mm : "0" + mm[0]) + (dd[1] ? dd : "0" + dd[0]); // padding
};

let LOCAL = false;
if (typeof AdsApp === 'undefined') {
  LOCAL = true;
}

function main() {
  console.log('Starting');
  let settings = scanForAccounts();
  let ids = Object.keys(settings);
  if (DEBUG) { ids = [""] }
  if (ids.length == 0) { console.log('No Rules Specified'); return; }
  if (isMCC()) {
    MccApp.accounts().withIds(ids).withLimit(50).executeInParallel('runRows', 'callBack', JSON.stringify(settings));
  } else {
    settings = settings[Object.keys(settings)[0]]
    for (let rowId in settings) {
      try {
        runScript(settings[rowId]);
      } catch (e) {
        console.error(e)
      }
    }
  }
}

/**
 * 
 * @param {object} settings 
 * @param {object} anomalySettings 
 */
function getEntities(settings, anomalySettings) {
  let afterDataEntities = getAfterEntityData(settings, anomalySettings);
  // console.log(`afterDataEntities: ${JSON.stringify(afterDataEntities)}`)
  if (Object.keys(afterDataEntities).length == 0) {
    return
  }

  let beforeDataEntities = getBeforeEntityData(settings, anomalySettings);
  // console.log(`beforeDataEntities: ${JSON.stringify(beforeDataEntities)}`);

  //now if we're at label level, we need to add up the campaigns which share the same label
  //product objects with the same makeup
  if (settings.REPORT_TYPE === reportTypes.Label) {
    beforeDataEntities = combineDataByLabel(beforeDataEntities, "beforeMetrics")
    afterDataEntities = combineDataByLabel(afterDataEntities, "afterMetrics")
  }

  let entities = combineBeforeAndAfterEntities(beforeDataEntities, afterDataEntities);

  return entities;

}

function getBeforeEntityData(settings, anomalySettings) {
  if (LOCAL) {
    const query = new Query(settings, anomalySettings, true).get();
    // console.log(query);
    return getBeforeEntityDummyData();
  }
  return getEntityData(settings, anomalySettings, true);
}

function getAfterEntityData(settings, anomalySettings) {
  if (LOCAL) {
    const query = new Query(settings, anomalySettings, false).get();
    // console.log(query);
    return getAfterEntityDummyData();
  }
  return getEntityData(settings, anomalySettings, false);
}

function combineBeforeAndAfterEntities(beforeDataEntities, afterDataEntities) {
  let entities = {}

  //pull everything into afterDataEntities object
  //add compare metricics (percententityIdes)
  for (let entityId in afterDataEntities) {
    if (typeof beforeDataEntities[entityId] === "undefined") {
      continue;
    }
    entities[entityId] = {};
    entities[entityId].beforeMetrics = beforeDataEntities[entityId].beforeMetrics;
    entities[entityId].afterMetrics = afterDataEntities[entityId].afterMetrics;
    entities[entityId].differences = {};
    for (let metric in afterDataEntities[entityId].afterMetrics) {
      const afterMetricValue = afterDataEntities[entityId].afterMetrics[metric];
      const beforeMetricValue = beforeDataEntities[entityId].beforeMetrics[metric];
      const difference = returnDifference(afterMetricValue, beforeMetricValue, "percent");
      entities[entityId].differences[metric] = difference;
      entities[entityId]["campaignName"] = afterDataEntities[entityId]["CampaignName"];
      entities[entityId]["adGroupName"] = afterDataEntities[entityId]["AdGroupName"];
    }

  }

  return entities;
}

/**
 * Get the entity (campaigns, ad groups or labels) data
 * as an object for before or after data
 * @param {object} settings 
 * @param {object} anomalySettings 
 * @param {boolean} compareRun 
 */
function getEntityData(settings, anomalySettings, compareRun) {
  const reportData = getReportDataFromReportType(settings.REPORT_TYPE);
  let id = reportData.id;//ID used in the object, will depend on the level
  const metricsKey = compareRun ? "beforeMetrics" : "afterMetrics";

  const query = new Query(settings, anomalySettings, compareRun).get();
  // console.log(`query: ${query}`);
  let report = AdsApp.report(query);
  let rows = report.rows();
  // console.log(`Number of rows: ${rows.totalNumEntities()}`);
  let entities = {};
  while (rows.hasNext()) {
    let row = rows.next();
    // console.log("Compare run? + query: " + compareRun + " - " + row.CampaignName);
    // console.log(`labels: ${row.Labels}`)
    if (row.Labels) {
      row.Labels = row.Labels.replace("[", "").replace("]", "").replace(/["']/g, "")
      row.Labels = row.Labels.split(",")
    }

    const entityId = settings.REPORT_TYPE === reportTypes.account ? "Account" : row[id];

    //for each alert param metric selected, store it
    for (let alertMetric_i in anomalySettings.alertMetrics) {
      let alertMetric = anomalySettings.alertMetrics[alertMetric_i];
      row[alertMetric] = convertMetricToNumber(row[alertMetric]);
      entities[entityId] = entities[entityId] || {};
      entities[entityId][metricsKey] = entities[entityId][metricsKey] || {};
      entities[entityId][metricsKey][alertMetric] = row[alertMetric];
      if (settings.REPORT_TYPE === reportTypes.AdGroup) {
        entities[entityId]["AdGroupName"] = row["AdGroupName"];
      }
      entities[entityId]["CampaignName"] = row["CampaignName"];
      entities[entityId]["Labels"] = row["Labels"];
    }
  }

  for (let name in entities) {
    if (typeof entities[name] == "undefined") {
      delete entities[name]
    }
  }

  return entities;
}

/**
 * If the metric is a string convert it to a number
 * @param {*} metric 
 */
function convertMetricToNumber(metric) {

  function percentageStringToFloat(metric) {
    metric = metric.replace(/%/g, '');
    return parseFloat(metric) / 100;
  }

  if (typeof metric === 'string' && metric.includes('%')) {
    return percentageStringToFloat(metric);
  }

  if (typeof metric === 'string') {
    metric = metric.replace(/,/g, '');
    return parseFloat(metric);
  }

  return metric;
}

/**
 * Filter the entities down to just the alerts which we'll report on
 * @param {object} entities 
 * @param {object} anomalySettings 
 */
function filterEntitiesToAlerts(entities, anomalySettings) {
  //decide what to keep based on alert params
  //  log("afterDataEntities: " + JSON.stringify(afterDataEntities))

  function getOperator(operator) {
    if (operator === ALERT_OPERATOR_ENUMS.INCREASE) {
      return ">"
    }
    if (operator === ALERT_OPERATOR_ENUMS.DECREASE) {
      return "<"
    }
  }

  /**
   * If decreasing the param needs to be a minus number
   * @param {number} param 
   * @param {string} operator
   */
  function getParam(param, operator) {
    if (operator === "<") {
      return -param;
    }
    return param;
  }

  function generateEvalString({ operator, param, difference }) {
    // console.log(`operator: ${operator}`)
    // console.log(`param: ${param}`)
    if (difference == "inf") {
      return `(${10000} ${operator} ${param})`
    }
    return `(${difference} ${operator} ${param})`
  }

  for (let id in entities) {
    let differences = entities[id]["differences"]
    for (let metrixIndex in anomalySettings.alertMetrics) {
      const metric = anomalySettings.alertMetrics[metrixIndex];
      const operator = getOperator(anomalySettings.alertOperators[metrixIndex]);
      const param = getParam(anomalySettings.alertParams[metrixIndex], operator);
      const difference = differences[metric];
      const evalString = generateEvalString({
        operator,
        param,
        difference,
      });

      if (!operator) {
        continue;
      }

      // log("evalString: " + evalString)
      if (!eval(evalString)) {
        delete entities[id]
      }
    }
  }
}

function runScript(settings) {
  log('Script Started');
  const usingDefaultSettings = settings.SETTINGS_URL === "";
  updateSettings(settings);
  addDefaultLog(settings)

  const settingsSpreadsheet = SpreadsheetApp.openByUrl(settings.SETTINGS_URL);
  addEditors(settingsSpreadsheet, settings.controlSheet, settings.LOGS_COLUMN);
  if (!isValidSettingsRow(settings)) {
    return;
  }

  log("Settings: " + JSON.stringify(settings));

  //don't process default settings (from the template sheet)
  if (usingDefaultSettings) {
    console.log("Using default settings. Please update the settings sheet and run the script again.")
    finaliseScript(settings);
    return;
  }

  const anomalySettings = populateAnomalySettings(settings);
  log("anomalySettings: " + JSON.stringify(anomalySettings))
  //have the information from the sheet, now to start grabbing data and making comparisons
  log("Calculating comparisons...")

  let entities = getEntities(settings, anomalySettings);
  // console.log(`entities before filter: ${JSON.stringify(entities)}`);

  // for (let id in entities) {
  //   console.log(`beforeMetrics: ${JSON.stringify(entities[id]['beforeMetrics'])}`)
  //   console.log(`afterMetrics: ${JSON.stringify(entities[id]['afterMetrics'])}`)
  // }
  console.log("Filtering down to alerts")
  filterEntitiesToAlerts(entities, anomalySettings);
  // console.log(`alert entities: ${JSON.stringify(entities)}`);


  const sheetArray = new SheetArray(settings, anomalySettings, entities).createSheetArray();
  if (LOCAL) {
    return;
  }
  if (!entities || Object.keys(entities).length == 0) {
    settings.LOGS.push("No alerts")
    finaliseScript(settings);
    return;
  }


  //create the sheet/folder if it doesn't already exist
  let reportsFolder = getFolder();
  // addEditors(reportsFolder, settings.controlSheet, settings.LOGS_COLUMN);
  addOutputSheetToSettings(settings, reportsFolder);

  //write to sheet
  log("Writing results...")
  writeMetricsToSheet(settings, sheetArray);

  //Send an email with the changes
  settings.htmlLogs = logsToHtml(settings.LOGS)

  if (shouldSend(settings, entities, 'email')) {
    emailSheet(settings);
  }
  if (shouldSend(settings, entities, 'slack')) {
    sendSlackMessage(settings);
  }

  finaliseScript(settings);


  //update settings/control sheet & throw the error if found
  //remove the unsuccessful note we added

  //*****************************************************   END OF MAIN   *****************************************************//
}


/**
 * 
 * @param {object} settings 
 * @param {object} entities 
 * @param {string} type email or slack
 * @returns {boolean} - Should the message be sent
 */
function shouldSend(settings, entities, type) {
  const sendKey = type === "email" ? "SEND_EMAIL" : "SEND_SLACK";
  if (!settings[sendKey]) {
    return false;
  }
  if (DEBUG) {
    return false;
  }
  if (Object.keys(entities).length === 0) {
    return false;
  }
  if (isWeekend() && !settings['WEEKEND_NOTIFICATIONS']) {
    return false;
  }
  const lastSentKey = type === "email" ? "EMAIL_LAST_SENT_TIME" : "SLACK_LAST_SENT_TIME";
  const lastSent = settings[lastSentKey];
  if (isWithinDays(lastSent, settings['NOTIFICATION_FREQUENCY'])) {
    return false;
  }
  return true;
}

function isWeekend() {
  const day = new Date().getDay();
  return day === 0 || day === 6;
}

/**
 * Defaults to false
 * @param {string} date 
 * @param {string} notificationFrequencyDays 
 * @returns boolean
 */
function isWithinDays(date, notificationFrequencyDays) {
  if (!date || !notificationFrequencyDays) {
    return false;
  }
  const notificationFrequencyDaysNumber = parseInt(notificationFrequencyDays);
  const lastSentDate = new Date(date);
  let cutOffDate = new Date();
  cutOffDate.setDate(cutOffDate.getDate() - notificationFrequencyDaysNumber);
  return lastSentDate >= cutOffDate;
}

function addOutputSheetToSettings(settings, reportsFolder) {
  const outputSheetName = 'Anomaly Detector - ' + settings.ALERT_NAME + ' - ' + AdsApp.currentAccount().getName();
  let outputSpreadsheet = getReportSpreadsheet(reportsFolder, outputSheetName, settings.OUTPUT_SHEET_URL);
  if (ANYONE_CAN_EDIT_SHEETS) {
    setSpreadsheetToAnyoneCanEdit(outputSpreadsheet);
  }
  settings.outputSheet = outputSpreadsheet.getActiveSheet();
  settings.OUTPUT_SHEET_URL = outputSpreadsheet.getUrl();
  log("Output: " + settings.OUTPUT_SHEET_URL);
  addEditors(outputSpreadsheet, settings.controlSheet, settings.LOGS_COLUMN);
}

function finaliseScript(settings) {
  settings.LOGS.push("The script ran sucessfully")
  settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN, 1, 1).setNote("")
  let errorMessage = ""
  updateControlSheet(errorMessage, settings)

  log('Finished');
}


/**
 * Build up the report query from the this.settings
 * That's:
 * 1) The main this.settings sheet this.settings
 * 2) The anomaly this.settings sheet this.settings
 */
class Query {

  constructor(settings, anomalySettings, compareRun) {
    this.settings = settings;
    this.anomalySettings = anomalySettings;
    this.compareRun = compareRun;
  }

  get() {
    const whereString = this._getWhereString();
    const duringString = this._getDuringString();
    const reportData = getReportDataFromReportType(this.settings.REPORT_TYPE);

    let selectString = "SELECT " + this.anomalySettings.allMetrics.join(',');
    if (reportData.columns.length) {
      selectString += ", " + reportData.columns.join(",");
    }
    let query = (selectString +
      " FROM  " + reportData.name +
      whereString +
      duringString);
    query = query.replace("EQUALS", "=")
    query = query.replace("NOT_EQUALS", "!=")
    // console.log(`compareRun: ${this.compareRun}`)
    // console.log(`query: ${query}`)
    return query;
  }

  _getDuringString() {
    const dates = this._getDates();
    if (this.compareRun) {
      return " DURING " + dates.compareStartDate + "," + dates.compareEndDate;
    }
    return " DURING " + dates.startDate + "," + dates.endDate;
  }

  /**
 * Turn days ago into dates (YYYYMMMDD)
 * @return {Object} - Object containing dates
 */
  _getDates() {
    let date = new Date();
    let preStartDate = new Date(date.getTime() - (this.settings.PREVIOUS_PERIOD_START * 24 * 60 * 60 * 1000));
    let preEndDate = new Date(date.getTime() - (this.settings.PREVIOUS_PERIOD_END * 24 * 60 * 60 * 1000));
    let postStartDate = new Date(date.getTime() - (this.settings.CURRENT_PERIOD_START * 24 * 60 * 60 * 1000));
    let postEndDate = new Date(date.getTime() - (this.settings.CURRENT_PERIOD_END * 24 * 60 * 60 * 1000));

    let compareStartDate = preStartDate.yyyymmdd();
    let compareEndDate = preEndDate.yyyymmdd();
    let startDate = postStartDate.yyyymmdd();
    let endDate = postEndDate.yyyymmdd();
    return {
      compareStartDate,
      compareEndDate,
      startDate,
      endDate
    }
  }

  _getWhereString() {
    let hasAttributes = this.anomalySettings.attributes.length > 0 ? true : false
    let hasthresholdMetricsPrevious = this.anomalySettings.thresholdMetricsPrevious.length > 0 ? true : false
    let hasthresholdMetricsCurrent = this.anomalySettings.thresholdMetricsCurrent.length > 0 ? true : false
    let whereString = "WHERE ";
    if (this.settings.REPORT_TYPE === reportTypes.AdGroup || this.settings.REPORT_TYPE === reportTypes.Campaign) {
      whereString += " CampaignStatus = ENABLED  "
    }
    if (this.settings.REPORT_TYPE === reportTypes.AdGroup) {
      whereString += " AND AdGroupStatus = ENABLED "
    }
    for (let i in this.anomalySettings.attributes) {
      whereString += this._getWhereOperator(whereString);
      whereString += this.anomalySettings.attributes[i] + " "
      whereString += this.anomalySettings.attributeOperators[i] + " '"
      whereString += this.anomalySettings.attributeParams[i] + "' "
      // whereString += (this.anomalySettings.attributes.length - 1).toFixed(0) > i ? " AND " : " "
    }

    if (hasthresholdMetricsCurrent && !this.compareRun) {
      for (let i in this.anomalySettings.thresholdMetricsCurrent) {
        whereString += this._getWhereOperator(whereString);
        whereString += this.anomalySettings.thresholdMetricsCurrent[i] + " "
        whereString += this.anomalySettings.thresholdOperatorsCurrent[i] + " "
        whereString += this.anomalySettings.thresholdParamsCurrent[i] + " "
        // whereString += (this.anomalySettings.thresholdMetricsCurrent.length - 1).toFixed(0) > i ? " AND " : " "
      }
    }

    if (hasthresholdMetricsPrevious && this.compareRun) {
      for (let i in this.anomalySettings.thresholdMetricsPrevious) {
        whereString += this._getWhereOperator(whereString);
        whereString += this.anomalySettings.thresholdMetricsPrevious[i] + " "
        whereString += this.anomalySettings.thresholdOperatorsPrevious[i] + " "
        whereString += this.anomalySettings.thresholdParamsPrevious[i] + " "
        // whereString += (this.anomalySettings.thresholdMetricsPrevious.length - 1).toFixed(0) > i ? " AND " : " "
      }
    }
    if (whereString.trim() === "WHERE") {
      whereString = "";
    }
    return whereString;
  }

  _getWhereOperator(whereString) {
    if (whereString.trim() === "WHERE") {
      return "";
    }
    return whereString.indexOf('WHERE') > -1 ? ' AND ' : '';
  }

}


/**
 * Populate the anomaly settings object from the Settings Sheet
 * @param {object} settings 
 */
function populateAnomalySettings(settings) {
  let anomalySettings = LOCAL ? getDummyAnomalySettings() : getSheetAnomalySettings(settings);
  if (anomalySettings.alertMetrics.length == 0) {
    anomalySettings.alertMetrics[0] = 'Clicks';
    anomalySettings.alertOperators[0] = ALERT_OPERATOR_ENUMS.INCREASE;
    anomalySettings.alertParams[0] = '-100';
  }
  return anomalySettings;
}

function getSheetAnomalySettings(settings) {
  let anomalySettings = {
    "thresholdMetricsPrevious": [],
    "thresholdOperatorsPrevious": [],
    "thresholdParamsPrevious": [],
    "thresholdMetricsCurrent": [],
    "thresholdOperatorsCurrent": [],
    "thresholdParamsCurrent": [],
    "attributes": [],
    "attributeOperators": [],
    "attributeParams": [],
    "alertMetrics": [],
    "alertParams": [],
    "alertOperators": [],
  }

  let moneyMetrics = ["TargetCpa", "CpvBid", "CpmBid", "CpcBid", "CostPerConversion", "CostPerAllConversion", "Cost", "AverageCpm", "AverageCpc", "AverageCost", "ActiveViewMeasurableCost", "ActiveViewCpm"]
  //threshold (filter) metrics - prev

  let settingsSpreadsheet = SpreadsheetApp.openByUrl(settings.SETTINGS_URL)
  settings.settingsSheet = settingsSpreadsheet.getActiveSheet()
  let row = 4;
  while (settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdMetricColumn_prev).getValue()) {
    let currentMetric = settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdMetricColumn_prev).getValue()
    let currentMetricParam = settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdParamColumn_prev).getValue()

    if (moneyMetrics.indexOf(currentMetric) > -1) {
      currentMetricParam = currentMetricParam * 1000000;
    }

    anomalySettings.thresholdMetricsPrevious.push(currentMetric);
    anomalySettings.thresholdOperatorsPrevious.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdOperatorColumn_prev).getValue());
    anomalySettings.thresholdParamsPrevious.push(currentMetricParam);
    row++;
  }
  //threshold (filter) metrics - curr
  row = 4;
  while (settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdMetricColumn_curr).getValue()) {
    let currentMetric = settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdMetricColumn_curr).getValue()
    let currentMetricParam = settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdParamColumn_curr).getValue()

    if (moneyMetrics.indexOf(currentMetric) > -1) {
      currentMetricParam = currentMetricParam * 1000000;
    }

    anomalySettings.thresholdMetricsCurrent.push(currentMetric);
    anomalySettings.thresholdOperatorsCurrent.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.thresholdOperatorColumn_curr).getValue());
    anomalySettings.thresholdParamsCurrent.push(currentMetricParam);
    row++;
  }
  //alert metrics
  row = 4;
  while (settings.settingsSheet.getRange(row, settingsColumnNumbers.alertMetricColumn).getValue()) {

    anomalySettings.alertMetrics.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.alertMetricColumn).getValue());
    anomalySettings.alertParams.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.alertParamColumn).getValue());
    anomalySettings.alertOperators.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.alertOperatorColumn).getValue());
    row++;
  }
  //attribute filters
  row = 4;
  while (settings.settingsSheet.getRange(row, settingsColumnNumbers.attributeColumn).getValue()) {
    anomalySettings.attributes.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.attributeColumn).getValue());
    anomalySettings.attributeParams.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.attributeParamColumn).getValue());
    anomalySettings.attributeOperators.push(settings.settingsSheet.getRange(row, settingsColumnNumbers.attributeOperatorColumn).getValue());
    row++;
  }
  let equalIndex = anomalySettings.attributeOperators.indexOf("EQUALS");
  anomalySettings.attributeOperators[equalIndex] = "=";
  equalIndex = anomalySettings.attributeOperators.indexOf("NOT_EQUALS");
  anomalySettings.attributeOperators[equalIndex] = "!=";

  anomalySettings.allMetrics = anomalySettings.thresholdMetricsPrevious.concat(anomalySettings.alertMetrics);
  anomalySettings.allMetrics = anomalySettings.allMetrics.concat(anomalySettings.thresholdMetricsCurrent);
  if (!anomalySettings.allMetrics.length) {
    anomalySettings.allMetrics = ["Impressions"]
  }

  return anomalySettings;
}


function writeMetricsToSheet(settings, sheetArray) {
  settings.outputSheet.clearContents();
  if (sheetArray.length == 0) {
    return;
  }
  settings.outputSheet.getRange(1, 1, sheetArray.length, sheetArray[0].length).setValues(sheetArray);
  //set first row to bold
  settings.outputSheet.getRange(1, 1, 1, sheetArray[0].length).setFontWeight("bold");

  settings.outputSheet.sort(3, false);
}

class SheetArray {
  constructor(settings, anomalySettings, entities) {
    this.settings = settings;
    this.anomalySettings = anomalySettings;
    this.entities = entities;
  }

  createSheetArray() {
    let sheetArray = [];
    const header = this.getHeader();
    sheetArray.push(header);
    this.addRows(sheetArray);
    return sheetArray;
  }

  addRows(sheetArray) {
    for (let entityId in this.entities) {
      let row = [];
      let entity = this.entities[entityId];
      if (this.settings.REPORT_TYPE === reportTypes.AdGroup || this.settings.REPORT_TYPE == reportTypes.Campaign) {
        row.push(entity.campaignName)
      }
      if (this.settings.REPORT_TYPE == reportTypes.AdGroup) {
        row.push(entity.adGroupName)
      }
      for (let metricIndex in this.anomalySettings.alertMetrics) {
        let metricName = this.anomalySettings.alertMetrics[metricIndex];
        let difference = entity.differences[metricName];
        difference = Math.round(difference * 10000) / 100 + "%"
        row.push(entity.beforeMetrics[metricName], entity.afterMetrics[metricName], difference);
      }
      sheetArray.push(row);
    }
  }

  getHeader() {
    let header = [];
    if (this.settings.REPORT_TYPE === reportTypes.AdGroup || this.settings.REPORT_TYPE === reportTypes.Campaign) {
      header.push("Campaign Name");
    }
    if (this.settings.REPORT_TYPE == reportTypes.AdGroup) {
      header.push("Ad Group Name");
    }
    for (let metricIndex in this.anomalySettings.alertMetrics) {
      let beforeString = this.anomalySettings.alertMetrics[metricIndex] + " (before)"
      if (header.indexOf(beforeString) > -1) { continue; }
      header.push(beforeString);
      header.push(this.anomalySettings.alertMetrics[metricIndex] + " (after)");
      header.push(this.anomalySettings.alertMetrics[metricIndex] + " % change");
    }
    return header;

  }
}



function returnFormat(type) {
  let numberFormat = ""
  if (type == "money") {
    numberFormat = "#,##0.00"
  } else if (type == "number") {
    numberFormat = "#,##0.00"
  } else if (type == "int") {
    numberFormat = "#,##0"
  } else if (type == "text") {
    numberFormat = ""
  } else if (type == "percentage") {
    numberFormat = "0.00%"
  } else if (type == "date") {
    numberFormat = "yyyy mmmm"
  } else {
    throw ("Error: number type not recognised, please check the header objects")
  }
  return numberFormat
}

/**
 * Get the report info for the AWQL query
 * @param {string} reportType 
 * @returns {Object} - Object containing report name, columns and id
 */
function getReportDataFromReportType(reportType) {
  let reportData = {};
  if (reportType === reportTypes.AdGroup) {
    reportData.name = " ADGROUP_PERFORMANCE_REPORT ";
    reportData.columns = ["AdGroupName", "CampaignName", "AdGroupId", "Labels"];
    reportData.id = "AdGroupId"
    return reportData;
  }
  if (reportType === reportTypes.Label) {
    reportData.name = " CAMPAIGN_PERFORMANCE_REPORT ";
    reportData.columns = ["CampaignName", "CampaignId", "Labels"];
    reportData.id = "CampaignId"
    return reportData;
  }
  if (reportType === reportTypes.Campaign) {
    reportData.name = " CAMPAIGN_PERFORMANCE_REPORT ";
    reportData.columns = ["CampaignName", "CampaignId"];
    reportData.id = "CampaignName"
    return reportData;
  }
  if (reportType === reportTypes.Account) {
    reportData.name = " ACCOUNT_PERFORMANCE_REPORT ";
    reportData.columns = [];
    reportData.id = ""
    return reportData;
  }
  throw new Error("Error, report type not recognised. Expected one of AdGroup, Label, Account or Campaign but got " + reportType)
}


function updateSettings(settings) {
  if (LOCAL) {
    return;
  }
  settings.controlSheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME)
  settings.LOGS_COLUMN = getLogsColumn(settings.controlSheet)
  settings.LOGS = [];
  settings.settingsSpreadsheet = getSettingsSpreadsheet(settings);
}

function setSpreadsheetToAnyoneCanEdit(spreadsheet) {
  const fileID = spreadsheet.getId();
  const file = DriveApp.getFileById(fileID);
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
}

/*
* Gets the report file (spreadsheet) for the given Google Ads account in the given folder.
* Creates a new spreadsheet if doesn't exist.
*/
function getSettingsSpreadsheet(settings) {
  const reportsFolder = getFolder();
  const sheetName = `Anomaly Detector Row Settings - ${settings.ALERT_NAME} (${AdsApp.currentAccount().getName()})`;
  if (!settings.SETTINGS_URL) {
    createSettingsSheet(settings, sheetName);
  }
  let spreadsheet = SpreadsheetApp.openByUrl(settings.SETTINGS_URL);
  spreadsheet.rename(sheetName);
  if (ANYONE_CAN_EDIT_SHEETS) {
    setSpreadsheetToAnyoneCanEdit(spreadsheet);
  }
  const file = DriveApp.getFilesByName(sheetName).next();
  try {
    reportsFolder.addFile(file);
  } catch (e) {
    console.error(e);
  }
  return spreadsheet;
}

function createSettingsSheet(settings, sheetName) {
  let spreadsheet = SpreadsheetApp.create(sheetName);
  settings.SETTINGS_URL = spreadsheet.getUrl();
  const templateSheet = SpreadsheetApp.openByUrl(SETTINGS_TEMPLATE_URL).getActiveSheet();
  templateSheet.copyTo(spreadsheet);
  spreadsheet.getSheetByName("copy of Anomaly Detector").setName("Anomaly Detector");
  spreadsheet.deleteSheet(spreadsheet.getSheetByName("Sheet1"))
  settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN - 2, 1, 1).setValue(settings.SETTINGS_URL);
  settings.LOGS = ["Settings sheet created. Add yours settings and run the script again."];
}

function addDefaultLog(settings) {
  if (LOCAL) {
    return;
  }
  let defaultNote = "Possible reasons include: 1) There was an error (check the logs within Google Ads) 2) The script was stopped before completion"
  settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN, 1, 1).setValue("The script did not run sucessfully").setNote(defaultNote)
}

function isValidSettingsRow(settings) {
  if (LOCAL) {
    return true;
  }
  if (!settings.FLAG) {
    let msg = "Run Script is not set to 'Yes'"
    settings.LOGS.push(msg)
    log(msg)
    updateControlSheet("", settings)
    return false;
  }

  if (settings.SETTINGS_URL == "") {
    let msg = "No settings sheet URL found"
    settings.LOGS.push(msg)
    updateControlSheet("", settings)
    return false;
  }
  return true;
}




function getHeaderTypes() {

  let headerTypes = {
    'ALERT_NAME': 'normal',
    'ID': 'normal',
    'NAME': 'normal',
    'EMAIL': 'csv',
    'FLAG': 'bool',
    'SUBJECT_PREFIX': 'normal',
    'SEND_EMAIL': 'bool',
    'EMAIL_LAST_SENT_TIME': 'normal',
    'SEND_SLACK': 'bool',
    'SLACK_LAST_SENT_TIME': 'normal',
    'NOTIFICATION_FREQUENCY': 'normal',
    'WEEKEND_NOTIFICATIONS': 'bool',
    'REPORT_TYPE': 'normal',
    'PREVIOUS_PERIOD_START': 'normal',
    'PREVIOUS_PERIOD_END': 'normal',
    'CURRENT_PERIOD_START': 'normal',
    'CURRENT_PERIOD_END': 'normal',
    'SETTINGS_URL': 'normal',
    'OUTPUT_SHEET_URL': 'normal'
  }
  return headerTypes;
}

function scanForAccounts() {
  log("getting settings...")
  let map = {};
  let controlSheet = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME)
  let data = SpreadsheetApp.openByUrl(INPUT_SHEET_URL).getSheetByName(INPUT_TAB_NAME).getDataRange().getValues();
  data.shift();
  data.shift();
  data.shift();

  log(JSON.stringify(data))
  //Run Script?	Level	Start Days Ago	End Days Ago	Start Days Ago	End Days Ago	Settings Sheet	Output
  let headerTypes = getHeaderTypes();

  let header = Object.keys(headerTypes)
  EMAIL_SENT_TIME_COLUMN_NUMBER = header.indexOf("EMAIL_LAST_SENT_TIME") + 1;
  SLACK_SENT_TIME_COLUMN_NUMBER = header.indexOf("SLACK_LAST_SENT_TIME") + 1;

  let LOGS_COLUMN = 0;
  let col = 5
  while (controlSheet.getRange(3, col).getValue()) {
    LOGS_COLUMN = controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) { break; }
    col++;
  }

  for (let k in data) {
    //if "run script" is not set to "yes", continue.
    if (data[k][header.indexOf("FLAG")] == '' || data[k][header.indexOf("FLAG")].toLowerCase() != 'yes') { continue; }
    let rowNum = parseInt(k, 10) + 4;
    let id = data[k][header.indexOf("ID")];
    let rowId = id + "/" + rowNum;
    map[id] = map[id] || {}
    map[id][rowId] = { 'ROW_NUM': (parseInt(k, 10) + 4) };
    for (let j in header) {
      if (header[j] == "LOGS_COLUMN") {
        map[id][rowId][header[j]] = LOGS_COLUMN;
        continue;
      }
      map[id][rowId][header[j]] = data[k][j];
    }
  }
  log(JSON.stringify(map))
  for (let id in map) {
    for (let rowId in map[id]) {
      for (let key in map[id][rowId]) {
        map[id][rowId][key] = processSetting(key, map[id][rowId][key], headerTypes, controlSheet)
      }
    }
  }
  log(JSON.stringify(map))
  return map;
}

function callBack() {
  // Do something here
  console.log('Finished');
}

function runRows(INPUT) {
  log("running rows")
  settings = JSON.parse(INPUT)[AdsApp.currentAccount().getCustomerId().toString()]
  for (let rowId in settings) {
    runScript(settings[rowId]);
  }
}

function processSetting(key, value, header, controlSheet) {
  let type = header[key]
  if (key == "ROW_NUM") {
    return value;
  }
  switch (type) {
    case "label":
      return [controlSheet.getRange(3, Object.keys(header).indexOf(key) + 1).getValue(), value]
      break;
    case "normal":
      return value
      break;
    case "bool":
      return value == "Yes" ? true : false;
      break;
    case "csv":
      let ret = value.split(",")
      ret = ret[0] == "" && ret.length == 1 ? [] : ret;
      if (ret.length == 0) {
        return [];
      } else {
        for (let r in ret) {
          ret[r] = String(ret[r]).trim()
        }
      }
      return ret;
      break;
    default:
      throw ("error setting type " + type + " not recognised for " + key)

  }
}




function updateControlSheet(errorMessage, settings) {
  let now = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), 'MMM dd, yyyy HH:mm:ss')
  if (errorMessage != "") {
    settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN, 1, 2).setValues([[errorMessage, now]])
    settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN, 1, 1).setNote("Note: Some rows running on account ID " + settings.ID + " may not have completed sucessfully. Please see their respective logs and 'Last Run' times.")
    throw (errorMessage)
  }

  //add final logs
  //stringify logs
  settings.LOGS = stringifyLogs(settings.LOGS)
  log("Logs: " + settings.LOGS)
  //update control sheet
  let put = [settings.OUTPUT_SHEET_URL, settings.LOGS, now]
  settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN - 1, 1, 3).setValues([put])
  settings.controlSheet.getRange(settings.ROW_NUM, settings.LOGS_COLUMN, 1, 1).setNote(settings.LOGS)
}

function logsToHtml(logs) {
  let html = "<ol>"
  for (let l in logs) {
    html += "<li>" + logs[l] + "</li>";
  }
  html += "</ol>"
  return html
}



/*
* Gets the report file (spreadsheet) for the given Google Ads account in the given folder.
* Creates a new spreadsheet if doesn't exist.
*/
function getReportSpreadsheet(folder, sheetName, url) {
  if (url != "" && typeof url != "undefined") {
    //we have a url, use that, but we may want to rename it if anything has changed
    let spreadsheet = SpreadsheetApp.openByUrl(url);
    spreadsheet.rename(sheetName)
    return spreadsheet;
  }

  let spreadsheet = SpreadsheetApp.create(sheetName);
  let file = DriveApp.getFileById(spreadsheet.getId());
  let oldFolder = file.getParents().next();
  folder.addFile(file);
  oldFolder.removeFile(file);

  return spreadsheet;
}
function getFolder() {
  if (FOLDER_ID) {
    return DriveApp.getFolderById(FOLDER_ID);
  }
  return DriveApp.getRootFolder();
}



function stringifyLogs(logs) {
  let s = ""
  for (let l in logs) {
    s += (parseInt(l) + 1) + ") ";
    s += logs[l] + " ";
  }
  return s
}

function round(num, n) {
  return +(Math.round(num + "e+" + n) + "e-" + n);
}

function getLogsColumn(controlSheet) {
  let col = 5
  let LOGS_COLUMN = 0;
  while (String(controlSheet.getRange(3, col).getValue())) {
    LOGS_COLUMN = controlSheet.getRange(3, col).getValue() == "Logs" ? col : 0;
    if (LOGS_COLUMN > 0) { break; }
    col++;
  }
  return LOGS_COLUMN;
}

function createTabs(tabNames, logSS) {
  //attempt to rename
  let logSheets = logSS.getSheets()
  for (let l in logSheets) {
    let logSheet = logSheets[l]
    try {
      logSheet.setName(tabNames[l])
    } catch (e) {

    }
  }
  //attempt to create
  for (let t in tabNames) {
    let tabName = tabNames[t]

    try {
      logSS.insertSheet(tabName)
    } catch (e) {

    }

  }

}

function addEditors(spreadsheet, controlSheet, LOGS_COLUMN) {

  //check current editors, add if they don't exist
  let currentEditors = spreadsheet.getEditors()
  let currentEditorEmails = []
  for (let c in currentEditors) {
    currentEditorEmails.push(currentEditors[c].getEmail().trim().toLowerCase());
  }

  let editors = controlSheet.getRange(1, LOGS_COLUMN).getValue().trim();
  if (editors == "") { return; }
  editors = editors.split(",").map((editor) => editor.trim().toLowerCase());

  for (let e in editors) {
    if (currentEditorEmails.indexOf(editors[e]) == -1) {
      spreadsheet.addEditor(editors[e])
    }
  }

}
function roundNumber(num, scale) {
  if (!("" + num).indexOf("e") > -1) {
    return +(Math.round(num + "e+" + scale) + "e-" + scale);
  } else {
    let arr = ("" + num).split("e");
    let sig = ""
    if (+arr[1] + scale > 0) {
      sig = "+";
    }
    return +(Math.round(+arr[0] + "e" + sig + (+arr[1] + scale)) + "e-" + scale);
  }
}


function settingsToTable(rowNumber, headerRow, lastColumn, controlSheet) {//settings, the row the headers are on, and the column the settings end at
  let columnTitles = controlSheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0]
  let settings = controlSheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0]
  let notes = controlSheet.getRange(headerRow, 1, 1, lastColumn).getNotes()[0]
  //create html table
  let table = "<table style='background-color:white;border-collapse:collapse;' border = 1 cellpadding = 5>  <tr>    <th>Setting</th>    <th>Value</th>     <th>Notes</th>  </tr>"
  for (let c in columnTitles) {
    table += "<tr>"
    table += "<td>" + columnTitles[c] + "</td>"
    table += "<td>" + settings[c] + "</td>"
    table += "<td>" + notes[c] + "</td>"
    table += "</tr>"
  }
  table += "</table>"
  return table
}

function emailSheet(settings) {
  log("Sending email")
  const accountName = AdsApp.currentAccount().getName();

  let subject = settings.SUBJECT_PREFIX + ' Anomaly Alert: ' + accountName + " - " + settings.ALERT_NAME + " - " + INPUT_TAB_NAME;
  console.log(`Email subject: ${subject}`);
  let msg = "<div style='line-height:1em;'><p>Hi,</p><p>You have alerts from the anomaly detector for the account '" + accountName + "'.</p><p>"
  if (settings.htmlLogs !== "") {
    msg += "<p>Script logs:</p>"
    msg += settings.htmlLogs
  }
  msg += "<br></br>"
  msg += "<p>The first 30 results are below, all of which are available on the <a href='" + settings.OUTPUT_SHEET_URL + "'>output sheet</a>:</p>"
  //create the table from the output
  let tab = SpreadsheetApp.openByUrl(settings.OUTPUT_SHEET_URL).getActiveSheet()
  let values = tab.getDataRange().getValues().slice(0, 30);

  msg += '<table style="background-color:white;border-collapse:collapse;" border = 1 cellpadding = 5>';
  for (let row = 0; row < values.length; ++row) {
    msg += "<tr>"
    for (let col = 0; col < values[0].length; ++col) {
      if (col < 2) {
        //first two columns are strings
        msg += isNaN(values[row][col]) || values[row][col] == "" ? '<td>' + values[row][col] + '</td>' : '<td>' + String(values[row][col]) + '</td>';
      } else {
        msg += isNaN(values[row][col]) || values[row][col] == "" ? '<td>' + values[row][col] + '</td>' : '<td>' + roundNumber(values[row][col], 2) + '</td>';
      }
    }
    msg += '</tr>';

  }
  msg += '</table><br><br>';
  msg += "<p>The settings for this row can be found on <a href='" + INPUT_SHEET_URL + "'>the control sheet</a> and below:</p><br>"
  msg += settingsToTable(settings.ROW_NUM, 3, settings.LOGS_COLUMN - 1, settings.controlSheet)
  msg += "</div>"
  log("email subject: " + subject)
  // log("email msg: " + msg)
  let emails = settings.EMAIL
  for (let email_i in emails) {
    log("Emailing '" + emails[email_i] + "'")
    MailApp.sendEmail({
      to: emails[email_i],
      subject: subject,
      htmlBody: msg
    });
  }

  settings.controlSheet.getRange(settings['ROW_NUM'], EMAIL_SENT_TIME_COLUMN_NUMBER).setValue(new Date());

}

function sendSlackMessage(settings) {
  const accountName = AdsApp.currentAccount().getName();
  let tab = SpreadsheetApp.openByUrl(settings.OUTPUT_SHEET_URL).getActiveSheet()
  let values = tab.getDataRange().getValues().slice(0, 30);
  const message = `Anomaly Alert:\n
    Name: ${settings.ALERT_NAME}\n
    Subject prefix: ${settings.SUBJECT_PREFIX}\n
    Account name: ${accountName}\n
    First 30 values: ${values}\n
    Logs: ${settings.LOGS}\n
    Control sheet: ${INPUT_SHEET_URL}\n
    Output sheet: ${settings.OUTPUT_SHEET_URL}\n
    `;
  new SlackNotification().sendToTeam(message);
  settings.controlSheet.getRange(settings['ROW_NUM'], SLACK_SENT_TIME_COLUMN_NUMBER).setValue(new Date());
}

/**
 * Aggregate the stats by label
 * @param {object} data 
 * @param {object} metrics 
 * @returns 
 */
function combineDataByLabel(data, metrics) {
  let labels = {}
  for (let id in data) {
    for (let labelIndex in data[id]["Labels"]) {
      let label = data[id]["Labels"][labelIndex];
      if (label === "--") { continue; }
      labels[label] = {}
    }
  }

  for (let id in data) {
    let metricData = data[id][metrics];
    for (let label in labels) {
      for (let metric in metricData) {
        labels[label][metrics] = labels[label][metrics] || {}
        labels[label][metrics][metric] = 0
      }
    }
  }

  for (let id in data) {
    let metricData = data[id][metrics];
    for (let label in labels) {
      for (let metric in metricData) {
        if (data[id]["Labels"].indexOf(label) > -1) {
          labels[label][metrics][metric] += parseFloat(metricData[metric], 10)
        }
      }
    }
  }
  return labels
}

function headersToSheet(settings, campaignName, adGroupName, alertMetrics) {
  log("adding headers to the sheet...")
  let headerValues = []
  if (settings.REPORT_TYPE === reportTypes.Campaign) {
    settings.outputSheet.getRange("A1:A1").setValue("Campaign");

  } else if (settings.REPORT_TYPE === reportTypes.Label) {
    settings.outputSheet.getRange("A1:A1").setValue("Label");

  } else if (settings.REPORT_TYPE == "AdGroup") {
    let headerValues = [[campaignName, adGroupName]];
    settings.outputSheet.getRange("A1:B1").setValues(headerValues);
    headerValues = []
  }

  for (let met_i in alertMetrics) {
    let beforeString = alertMetrics[met_i] + " (before)"
    // log(beforeString, headerValues.indexOf(beforeString), headerValues.indexOf(beforeString)>-1)
    if (headerValues.indexOf(beforeString) > -1) { continue; }
    headerValues.push(beforeString);
    headerValues.push(alertMetrics[met_i] + " (after)");
    headerValues.push(alertMetrics[met_i] + " % change");
  }

  let metricHeaders = []; metricHeaders.push(headerValues);
  let rowStart = settings.REPORT_TYPE === reportTypes.Campaign || settings.REPORT_TYPE === reportTypes.Label ? 2 : 3
  settings.outputSheet.getRange(1, rowStart, 1, metricHeaders[0].length).setValues(metricHeaders);
}

function compare(metric, operator, compareMetric) {
  if (operator == ">") {
    if (metric > compareMetric) { return true } else { return false }
  } else if (operator == "<") {
    if (metric < compareMetric) { return true } else { return false }
  }
}

function returnDifference(number, compareNumber, type) {
  //needs to return 100% if that's the case
  //needs to return infitiy if that's the case
  if (number == 0 && compareNumber == 0) { return 0; }
  if (number == "undefined" || compareNumber == "undefined" || number == "NaN" || compareNumber == "NaN") {
    return "undefined";
  } else {
    if (type == "number") {
      return parseFloat(number - compareNumber);
    } else if (type == "percent") {
      return (parseFloat(number) > 0 && parseFloat(compareNumber) == 0) ? "inf" : parseFloat((number - compareNumber) / compareNumber);
    } else {
      return "error, expected either 'percent' or 'number' as type";
    }
  }
}

function decimalAdjust(type, value, exp) {
  // If the exp is undefined or zero...
  if (typeof exp === 'undefined' || +exp === 0) {
    return Math[type](value);
  }
  value = +value;
  exp = +exp;
  // If the value is not a number or the exp is not an integer...
  if (isNaN(value) || !(typeof exp === 'number' && exp % 1 === 0)) {
    return NaN;
  }
  // If the value is negative...
  if (value < 0) {
    return -decimalAdjust(type, -value, exp);
  }
  // Shift
  value = value.toString().split('e');
  value = Math[type](+(value[0] + 'e' + (value[1] ? (+value[1] - exp) : -exp)));
  // Shift back
  value = value.toString().split('e');
  return +(value[0] + 'e' + (value[1] ? (+value[1] + exp) : exp));
}

if (!Math.round10) {
  Math.round10 = function (value, exp) {
    return decimalAdjust('round', value, exp);
  };
}


function getHtmlTable(range) {
  let ss = range.getSheet().getParent();
  let sheet = range.getSheet();
  startRow = range.getRow();
  startCol = range.getColumn();
  lastRow = range.getLastRow();
  lastCol = range.getLastColumn();

  // Read table contents
  let data = range.getValues();

  // Get css style attributes from range
  let fontColors = range.getFontColors();
  let backgrounds = range.getBackgrounds();
  let fontFamilies = range.getFontFamilies();
  let fontSizes = range.getFontSizes();
  let fontLines = range.getFontLines();
  let fontWeights = range.getFontWeights();
  let horizontalAlignments = range.getHorizontalAlignments();
  let verticalAlignments = range.getVerticalAlignments();

  // Get column widths in pixels
  let colWidths = [];
  for (let col = startCol; col <= lastCol; col++) {
    colWidths.push(sheet.getColumnWidth(col));
  }
  // Get Row heights in pixels
  let rowHeights = [];
  for (let row = startRow; row <= lastRow; row++) {
    rowHeights.push(sheet.getRowHeight(row));
  }

  // Future consideration...
  let numberFormats = range.getNumberFormats();

  // Build HTML Table, with inline styling for each cell
  let tableFormat = 'style="border:1px solid black;border-collapse:collapse;text-align:center" border = 1 cellpadding = 5';
  let html = ['<table ' + tableFormat + '>'];
  // Column widths appear outside of table rows
  for (col = 0; col < colWidths.length; col++) {
    html.push('<col width="' + colWidths[col] + '">')
  }
  // Populate rows
  for (row = 0; row < data.length; row++) {
    html.push('<tr height="' + rowHeights[row] + '">');
    for (col = 0; col < data[row].length; col++) {
      // Get formatted data
      let cellText = data[row][col];
      if (cellText instanceof Date) {
        cellText = Utilities.formatDate(
          cellText,
          ss.getSpreadsheetTimeZone(),
          'MMM/d EEE');
      }
      let style = 'style="'
        + 'color: ' + fontColors[row][col] + '; '
        + 'font-family: ' + fontFamilies[row][col] + '; '
        + 'font-size: ' + fontSizes[row][col] + '; '
        + 'font-weight: ' + fontWeights[row][col] + '; '
        + 'background-color: ' + backgrounds[row][col] + '; '
        + 'text-align: ' + horizontalAlignments[row][col] + '; '
        + 'vertical-align: ' + verticalAlignments[row][col] + '; '
        + '"';
      html.push('<td ' + style + '>'
        + cellText
        + '</td>');
    }
    html.push('</tr>');
  }
  html.push('</table>');

  return html.join('');
}

let log = function () {
  let message = ""
  for (let i = 0; i < arguments.length; i++) {
    message += arguments[i];
    if (i < arguments.length - 1) message += ", "
  }
  if (LOCAL) {
    console.log(message)
    return;
  }
  console.log(AdsApp.currentAccount().getName() + ": " + message)
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

function getBeforeEntityDummyData() {
  return {}
}

function getAfterEntityDummyData() {
  return {}
}

function getDummyAnomalySettings() {
  return {
    "thresholdMetricsPrevious": [],
    "thresholdOperatorsPrevious": [],
    "thresholdParamsPrevious": [],
    "thresholdMetricsCurrent": [],
    "thresholdOperatorsCurrent": [],
    "thresholdParamsCurrent": [],
    "attributes": [],
    "attributeOperators": [],
    "attributeParams": [],
    "alertMetrics": [
    ],
    "alertParams": [
    ],
    "alertOperators": [
    ],
    "allMetrics": [
      "Clicks"
    ]
  }
}

if (LOCAL) {
  const settings = { "ROW_NUM": 4, "ALERT_NAME": "Google Ads Down (Yesterday)", "ID": "111-032-7969", "NAME": "KP", "EMAIL": "charles@adbloomdigital.com, symonds@kitchenprovisions.co.uk", "FLAG": "Yes", "SUBJECT_PREFIX": "Urgent - Google Ads is Down!", "SEND_EMAIL": "Yes", "EMAIL_LAST_SENT_TIME": "2024-07-08T22:41:39.022Z", "SEND_SLACK": "Yes", "SLACK_LAST_SENT_TIME": "2024-07-08T22:14:11.951Z", "NOTIFICATION_FREQUENCY": 2, "REPORT_TYPE": "Account", "PREVIOUS_PERIOD_START": 14, "PREVIOUS_PERIOD_END": 14, "CURRENT_PERIOD_START": 1, "CURRENT_PERIOD_END": 1, "SETTINGS_URL": "https://docs.google.com/spreadsheets/d/16jqLAPGr9yRL6zT-TtuFcvc5IlBAIVfAHsCv_HLZq8w/edit", "OUTPUT_SHEET_URL": "https://docs.google.com/spreadsheets/d/1WJUigdP4XhIUziLKvatuo78w4nbILGpJ5sAg873me5g/edit" }
  runScript(settings);
}





class SlackNotification {

  constructor() {
    try {
      this.teamWebhookUrl = SLACK_TEAM_WEBHOOK_URL;
      this.adminWebhookUrl = SLACK_ADMIN_WEBHOOK_URL;

      if (typeof SENT_SLACK_NOTIFICATIONS === "undefined") {
        this.sendToAdmin(`SENT_SLACK_NOTIFICATIONS array not defined for '${AdsApp.getExecutionInfo().getScriptName()}' script `);
      }

    } catch (e) {
      this._logError(e);
    }
  }

  sendToTeam(text) {
    try {
      this._sendTeamMessage(text);
    } catch (e) {
      this._logError(e);
      this.sendToAdmin(`Error sending slack message to team: ${e}`);
    }
  }

  sendToAdmin(text) {
    try {
      this._sendAdminMessage(text);
    } catch (e) {
      this._logError(e);
    }
  }

  _sendTeamMessage(text) {
    if (APP_ENVIRONMENT !== APP_ENVIRONMENTS.PRODUCTION) {
      console.warn(`Team notifications won't be sent outside of the production environment`)
      return;
    }
    if (!this.teamWebhookUrl) {
      return;
    }
    this._send(text, this.teamWebhookUrl);
  }

  _sendAdminMessage(text) {
    if (typeof APP_ENVIRONMENT !== "undefined" && APP_ENVIRONMENT !== APP_ENVIRONMENTS.PRODUCTION) {
      text = `*Sending from ${APP_ENVIRONMENT}*\n ${text}`
    }
    this._send(text, this.adminWebhookUrl);
  }

  _send(text, webhookUrl) {
    const scriptName = AdsApp.getExecutionInfo().getScriptName();
    text = `Message from the '${scriptName}' script:\n${text}`;
    try {
      this._sendSlackMessage(text, webhookUrl);
    } catch (e) {
      this._logError(e);
    }
  }

  _sendSlackMessage(text, webhookUrl) {
    if (SENT_SLACK_NOTIFICATIONS.includes(text)) {
      return;
    }
    const slackMessage = {
      text: text,
      icon_url:
        'https://www.gstatic.com/images/icons/material/product/1x/adwords_64dp.png',
      username: 'Google Ads Scripts',
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(slackMessage)
    };
    UrlFetchApp.fetch(webhookUrl, options);

    SENT_SLACK_NOTIFICATIONS.push(text);
  }

  _logError(error) {
    console.warn('Could not send slack notification. The script logic will not be affected. Here is the error:')
    console.warn(error);
  }

}