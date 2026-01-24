/**
* Auto-negative keyword adder for Shabba.io
* Pro Version
* @author Charles Bannister
* Connect with me on LinkedIn: https://www.linkedin.com/in/charles-bannister-92a1a228/
*
* This script will automatically add negative keywords based on specified criteria
* 
* Note updating the version (the variable) will update the sheet
* Version: 4.2.0
*
*Updates:
- 1.8 - converted positive keywords to lowercase making then case-insensitive
- 1.9 - added pre-emptive negative keyword functionality
- 2.0 - script allows multiple campaign names for each sheet
- 2.1 - column indexing changed to accommodate new columns for traffic volume features
- 2.2 - iterate over adgroups, then campaigns
- 2.3 - added contains/not contains query options
- 2.3.1 - added full list error email
- 2.4 - If lists are full, create a new list
    -2.5 - allow comma separated positive keywords
- 2.5 - added ignore spaces option
- 2.6 - added ad group name as keywords functionality
- 2.7 - moved settings including query contains for a customer
- 2.8 - added [exact match] positive keyword functionality
- 2.8.1 - .toLowerCase() on non-string fix
- 2.9.0 - updated to run through ad groups based on their timestamp (instead of sheets)
* 3.0.0:
*  - added multi campaign support in
*  - reworked new negative keyword list creation if full
*  - added preview mode
*  - added campaign name and ad group name pattern matching e.g. regex
*  - added regex keywords
*  - added email alerts
* 3.1.0 - combined the single account & MCC versions (different templates)
* 4.0.0 - added approx match, campaign level support, output sheets and a new layout
**/

let INPUT_SHEET_URL = "YOUR_SPREADSHEET_URL_HERE";


//You may also be interested in this Chrome keyword wrapper!
//https://chrome.google.com/webstore/detail/keyword-wrapper/paaonoglkfolneaaopehamdadalgehbb

//Template: https://docs.google.com/spreadsheets/d/1vJyQ9PT8u7pMZ4X5WoJf4SnxWt_KfRIfOfI7Q6BWXno



// editing below this line is encouraged, but you might want to create a backup first

//DELETE ABOVE THIS LINE BEFORE SCRAMBLING
const TEST_MODE = false;

//When updating the version number
//Also update the SheetAdmin with the number of columns, etc.
const VERSION_NUMBER = '4.2.0';
const VERSION_NUMBER_RANGE = 'B1'

const LOG_INVALID_SEARCH_TERMS = true;

const SCRIPT_NAME = isMCC() ? 'Auto Negative Keywords Pro (MCC)' : 'Auto Negative Keywords Pro';
const SHABBA_SCRIPT_ID = isMCC() ? 5 : 1;

const OUTPUT_SHEET_NAME = 'Output';

let SPREADSHEET = SpreadsheetApp.openByUrl(INPUT_SHEET_URL);
const STATUSES = {
  "ENABLED": "ENABLED",
  "PAUSED": "PAUSED",
  "NOT_FOUND": "NOT_FOUND",
}

const STATUS_ENUMS = {
  "ENABLED": "Enabled",
  "PAUSED": "Paused",
  "NOT_FOUND": "Not found",
}

const AD_GROUP_COMMANDS = {
  'RUN_ALL_ENABLED': 'RUN_ALL_ENABLED',
  'RUN_FIRST_ENABLED': 'RUN_FIRST_ENABLED',
}

const ENTITY_TYPES = {
  'CAMPAIGN': 'CAMPAIGN',
  'AD_GROUP': 'AD_GROUP',
  'KEYWORD': 'KEYWORD',
}

const ADD_NEGATIVE_TO_OPTIONS = {
  'LIST': 'List (Add Name)',
  'AD_GROUP': 'Ad Group',
  'CAMPAIGN': 'Campaign',
}

const SETTING_CELL_REFERENCES = {
  ACCOUNT_ID: { reference: 'A4', type: 'default' },
  CLICKS_MORE_THAN: { reference: 'B4', type: 'default' },
  IMPRESSIONS_MORE_THAN: { reference: 'C4', type: 'default' },
  CONVERSIONS_LESS_THAN: { reference: 'D4', type: 'default' },
  DATE_RANGE: { reference: 'E4', type: 'default' },
  PULL_FROM: { reference: 'F4', type: 'default' },
  CONTAINS: { reference: 'G4', type: 'csv' },
  NOT_CONTAINS: { reference: 'H4', type: 'csv' },
  MAX_SEARCH_TERMS: { reference: 'I4', type: 'default' },
  EMAILS: { reference: 'A7', type: 'csv' },
  EMAIL_ALERT: { reference: 'B7', type: 'boolean' },
  PREVIEW_MODE: { reference: 'C7', type: 'boolean' },
  NEW_NEGATIVES_ONLY: { reference: 'D7', type: 'boolean' },
  RUN_KEYWORDLESS_AD_GROUPS: { reference: 'E7', type: 'boolean' },
  AD_GROUP_NAMES_AS_KEYWORDS: { reference: 'F7', type: 'default' },
  CREATE_LIST_IF_FULL: { reference: 'G7', type: 'boolean' },
  NEGATIVE_MATCH_TYPE: { reference: 'H7', type: 'default' },
  MIN_WORDS: { reference: 'I7', type: 'default' },
  MAX_WORDS: { reference: 'J7', type: 'default' },
  LLM: { reference: 'J4', type: 'default' },
  API_KEY: { reference: 'K4', type: 'default' },
  MOCK_RESPONSE: { reference: 'L4', type: 'default' },
  SKIP_PREVIOUSLY_CHECKED: { reference: 'M4', type: 'boolean' },
}



const TIMESTAMP_COLUMN = 3;
const EMAIL_TIMESTAMP_COLUMN = 4;

const ACCOUNT_ID_CELL_REFERENCE = 'A4';

const ENTITY_NAME_CELL_REFERENCES = ['D13', 'D14', 'D15', 'D16'];

const ROW_NUMBERS = {
  LOGS: 9,
  LAST_RUN: 10,
  SKIP: 11,
  LOG_TO_SHEET_BOOL: 12,
  OUTPUT_SHEET_URL: 13,
  NUM_MATCHES: 14,
  CAMPAIGN_NAME: 15,
  CAMPAIGN_ID: 16,
  AD_GROUP_NAME: 17,
  AD_GROUP_ID: 18,
  KEYWORD_LIST_NAME: 19,
  ADD_NEGATIVE_TO: 20,
  APPROX_MATCH: 21,
  APPROX_MATCH_THRESHOLD: 22,
  FIRST_POSITIVE_KEYWORD: 23,
};

const FIRST_RULE_COLUMN_NUMBER = 4;

const RULE_PARAM_KEYS = {
  'skip': 'SKIP',
  'logToSheetBool': 'LOG_TO_SHEET_BOOL',
  'outputSheetUrl': 'OUTPUT_SHEET_URL',
  'numberOfMatches': 'NUM_MATCHES',
  'campaignName': 'CAMPAIGN_NAME',
  'campaignId': 'CAMPAIGN_ID',
  'adGroupName': 'AD_GROUP_NAME',
  'adGroupId': 'AD_GROUP_ID',
  'negativeListName': 'KEYWORD_LIST_NAME',
  'addNegativeTo': 'ADD_NEGATIVE_TO',
  'approxMatch': 'APPROX_MATCH',
  'approxMatchThreshold': 'APPROX_MATCH_THRESHOLD',
}
const campaignTypes = {
  "shopping": "Shopping",
  "search": "Search",
}

const GLOBAL_FILTERS = [
  { field: 'campaign.experiment_type', operator: '=', value: 'BASE' },
  { field: 'campaign.status', operator: '=', value: 'ENABLED' },
  { field: 'campaign.advertising_channel_sub_type', operator: 'NOT IN', value: "(SEARCH_EXPRESS, SEARCH_MOBILE_APP)" },
];

// console.log("Thanks for supporting Shabba.io");




function main() {
  if (TEST_MODE) {
    console.warn('Test mode is enabled. \nTest logs will be written to the console. \nTest logs will slow the script down considerably and may cause timeouts.\n\n');
  }
  console.log('Started');
  if (!isMCC()) {
    runAccount();
    return;
  }
  let ids = getAccountIds();
  console.log('Account IDs: ' + ids);
  MccApp.accounts()
    .withIds(ids)
    .withLimit(50)
    .executeInParallel("runAccount");
}


function getAccountIds() {
  let sheets = SPREADSHEET.getSheets();
  let accountIds = [];
  for (let sheet of sheets) {
    if (
      sheet
        .getName()
        .toLowerCase()
        .indexOf("(skip)") > -1
    )
      continue;
    let accountId = sheet
      .getRange(ACCOUNT_ID_CELL_REFERENCE)
      .getValue()
      .trim();
    if (!accountIdIsValid(accountId)) continue;
    if (accountIds.indexOf(accountId) > -1) continue; //already added
    accountIds.push(accountId);
  }
  if (accountIds.length == 0) {
    console.log("No valid account IDs were found in the spreadsheet. Please check IDs follow the 000-000-0000 format.")
  }
  return accountIds;
}

function accountIdIsValid(accountId) {
  //each account id should have 2 dashes and 10 numbers
  let split = accountId.split("");
  if (split.length !== 12) return false;
  split = accountId.split("-");
  if (split.length !== 3) return false;
  return true;
}

/**
 * Run all Sheets for a single account
 */
function runAccount() {
  new AddAdGroupData().add();
  const runTimes = new GetRunTimes(ROW_NUMBERS).getLastRunTimes();
  let sortedSheetNames = runTimes.sort((a, b) => a.lastRunTime - b.lastRunTime).map((item) => item.sheetName);
  let uniqueSheetNames = [...new Set(sortedSheetNames)];
  testLog('uniqueSheetNames: ' + uniqueSheetNames);
  let negativeList = new NegativeList();
  negativeList.createNegativeListsMap();
  testLog(`Negative lists: ${Object.keys(negativeList.negativeListsMap)}`);
  for (let sheetName of uniqueSheetNames) {

    //MCC logic - only run this account's sheets
    let sheet = SPREADSHEET.getSheetByName(sheetName);
    const sheetAccountId = sheet.getRange(ACCOUNT_ID_CELL_REFERENCE).getValue();
    testLog('sheetAccountId: ' + sheetAccountId);
    testLog('Customer ID: ' + AdsApp.currentAccount().getCustomerId());
    if (isMCC() && sheetAccountId !== AdsApp.currentAccount().getCustomerId()) continue;
    runSheet(sheetName, negativeList, sheet);
  }

}
class SheetAdmin {

  constructor(sheet) {
    this.sheet = sheet;
  }

  update() {
    if (!this._isRulesSheet()) {
      testLog(this.sheet.getName() + " is not a rules sheet")
      return;
    }
    if (this._hasLatestVersion()) {
      testLog("Already the latest version")
      return;
    }
    this._addScriptLink();
    this._updateVersion();
  }

  _addScriptLink() {
    const range = this.sheet.getRange(1, 1);
    const formula = `=HYPERLINK("https://shabba.io/script/${SHABBA_SCRIPT_ID}", "View the script page for support and info")`;
    range.setFormula(formula);
  }

  _isRulesSheet() {
    return this.sheet.getRange("A3").getValue() === "Account ID (MCC Only)";
  }

  _hasLatestVersion() {
    return this.sheet.getRange(VERSION_NUMBER_RANGE).getValue().replace('Version: ', '') === VERSION_NUMBER;
  }

  _updateVersion() {
    this.sheet.getRange(VERSION_NUMBER_RANGE).setValue('Version: ' + VERSION_NUMBER);
  }



}

class CampaignInfo {

  getCampaignStatusFromId(campaignId) {
    if (!campaignId) {
      return STATUSES.NOT_FOUND;
    }
    let cols = ["campaign.status"];
    let reportName = "campaign";
    let where = `where campaign.id = '${campaignId}'`;
    let query = ["select", cols.join(","), "from", reportName, where].join(" ");
    let reportIter = AdsApp.report(query).rows();

    if (reportIter.hasNext()) {
      let row = reportIter.next();
      return row['campaign.status'];
    }
    return STATUSES.NOT_FOUND;
  }

  _campaignExists() {
    if (this.campaignData.status === STATUSES.NOT_FOUND) {
      return false;
    }
    return true;
  }

  _getCampaignDataFromName(campaignName) {
    if (campaignName.trim() == "") {
      // log("Can't get status campaign name is " + campaignName)
      return {};
    }
    let cols = ["CampaignStatus", "CampaignId"];
    let reportName = "CAMPAIGN_PERFORMANCE_REPORT";
    let where = "where CampaignName = '" + campaignName.trim() + "'";
    let query = ["select", cols.join(","), "from", reportName, where].join(" ");
    let reportIter = AdsApp.report(query).rows();

    if (reportIter.hasNext()) {
      let row = reportIter.next();
      return { status: row.CampaignStatus.toUpperCase(), id: row.CampaignId };
    }
    return { status: STATUSES.NOT_FOUND, id: "" };
  }



}



class StringPattern {

  constructor() {
    this.sheetStringFunctions = {
      'regex': 'REGEXP_MATCH',
      'not_regex': 'NOT REGEXP_MATCH',
      'contains': 'LIKE',
      'not_contains': 'NOT LIKE',
      'like': 'LIKE',
      'not_like': 'NOT LIKE',
      'equals': '=',
    }
  }

  operatorValueStringFromPattern(patternString) {
    for (const functionName in this.sheetStringFunctions) {
      if (!patternString.startsWith(functionName)) {
        continue;
      }
      let value = patternString.substring(functionName.length + 1, patternString.length - 1);
      if (functionName.includes('contains')) {
        value = `%${value}%`;
      }
      return `${this.sheetStringFunctions[functionName]} '${value}'`;
    }
    return `= '${patternString}'`; // default to equals
  }

}

class EntityFromPattern {

  constructor(type, pattern, campaignId = '') {
    this.type_enums = { 'AD_GROUP': 'Ad Group', 'CAMPAIGN': 'Campaign' };
    this.type = type;
    this.pattern = pattern || '';
    this.campaignId = campaignId;
    if (this.type === 'AD_GROUP' && this.campaignId === '') {
      throw new Error('Campaign name or ID is required for AD_GROUP type');
    }
  }

  getEntityId() {
    const query = this._getQuery();
    testLog(`Query: ${query}`);
    const report = AdsApp.report(query);
    const rows = report.rows();
    if (rows.hasNext()) {
      const row = rows.next();
      testLog(row[`${this.type.toLowerCase()}.id`]);
      return row[`${this.type.toLowerCase()}.id`];
    }
    console.warn(`The pattern ${this.pattern} did not match any ${this.type_enums[this.type]}s in the account.`);
  }

  _getQuery() {
    const stringPattern = new StringPattern();
    const operatorValueString = stringPattern.operatorValueStringFromPattern(this.pattern);
    const type = this.type.toLowerCase();
    let query = `SELECT ${type}.id from ${type} where ${type}.name ${operatorValueString}`;
    query += ' and campaign.status = "ENABLED"';
    if (this.type === 'AD_GROUP') {
      query += ' and ad_group.status = "ENABLED"';
    }
    if (this.type === 'AD_GROUP') {
      query += ` and campaign.id = '${this.campaignId}' `
    }
    return query;
  }

  isValidPattern() {
    testLog(`this.pattern: ${this.pattern}`)
    if (!this.pattern || typeof this.pattern !== 'string' || this.pattern.trim() === '') {
      return false;
    }
    const validStartStrings = Object.keys(new StringPattern().sheetStringFunctions);
    for (const startString of validStartStrings) {
      if (this.pattern.startsWith(startString)) {
        return true;
      }
    }
    return false;
  }

}



/**
* Get AdWords Formatted date for n days back
* @param {int} d - Numer of days to go back for start/end date
* @return {String} - Formatted date yyyyMMdd
**/
function getAdWordsFormattedDate(d, format) {
  var date = new Date();
  if (d) {
    date.setDate(date.getDate() - d);
  }
  return Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), format);
}


/**
 * This is for getting the run times before the script runs
 * We'll grab all of the last run times from all the sheets
 * Then decide which sheet to run first based on the earliest time
 * Separate logic will then run the earliest Ad Group within the selected sheet
 */
class GetRunTimes {

  constructor(rowNumbers) {
    this.rowNumbers = rowNumbers;
  }

  /**
   * Get the last run time for each ad group
   * from all of the sheets
   * 
   */
  getLastRunTimes() {
    let runTimes = [];
    let sheets = SPREADSHEET.getSheets();
    for (let sheet of sheets) {
      let sheetName = sheet.getName();
      if (sheetName.indexOf("(skip)") > -1 || sheetName.indexOf("Ad Group Data") > -1) {
        continue
      };
      runTimes = runTimes.concat(this.getSheetLastRunTimes(sheet))
    }
    return runTimes;
  }

  /**
   * Get the last run time for each ad group for a single sheet
   * return the array sorted by earliest to latest
   * @param {SpreadsheetApp.sheet} sheet 
   * @returns {Array} runTimes
   */
  getSheetLastRunTimes(sheet) {
    let runTimes = [];
    let columnNumber = FIRST_RULE_COLUMN_NUMBER;
    while (sheet.getRange(this.rowNumbers.NUM_MATCHES, columnNumber).getValue()) {
      let lastRunTime = this._getAdGroupLastRunTime(columnNumber, sheet);
      if (!this._includeRule(columnNumber, sheet)) {
        columnNumber++;
        continue;
      }
      const sheetName = sheet.getName();
      const adGroupName = sheet.getRange(this.rowNumbers.AD_GROUP_NAME, columnNumber).getValue();
      runTimes.push({ sheetName, adGroupName, lastRunTime, columnNumber })
      columnNumber++;
    }
    runTimes = runTimes.sort((a, b) => a.lastRunTime - b.lastRunTime);
    return runTimes
  }

  /**
   * Whether the rule should be considered/included
   * @param {number} column
   * @param {SpreadsheetApp.sheet} sheet
   * @returns {boolean}
   * 
   */
  _includeRule(column, sheet) {
    let skipAdGroup = sheet.getRange(this.rowNumbers.SKIP, column).getValue();
    if (skipAdGroup) {
      return false;
    }
    let campaignId = sheet.getRange(this.rowNumbers.CAMPAIGN_ID, column).getValue();
    let campaignName = sheet.getRange(this.rowNumbers.CAMPAIGN_NAME, column).getValue();
    if (!campaignId && !campaignName) {
      return false;
    }
    let entityName;
    for (let cellRef of ENTITY_NAME_CELL_REFERENCES) {
      let value = sheet.getRange(cellRef).getValue();
      if (value) {
        entityName = value;
        break;
      }
    }
    if (!entityName) {
      return false;
    }
    return true;
  }

  /**
   * Get the last run time from a single ad group
   * If the ad group has never run before, return a date in the past
   * @param {number} column 
   * @param {SpreadsheetApp.sheet} sheet 
   * @returns {Date} lastRunTime
   */
  _getAdGroupLastRunTime(column, sheet) {
    let lastRunTime = sheet.getRange(this.rowNumbers.LAST_RUN, column).getValue();
    if (!lastRunTime) {
      lastRunTime = new Date(0)
    }
    return lastRunTime;
  }

}

function setColumnLog(sheet, columnNumber, logText, noteText) {
  sheet.getRange(ROW_NUMBERS.LOGS, columnNumber).setValue(logText);
  sheet.getRange(ROW_NUMBERS.LOGS, columnNumber).setNote(noteText);
}

function setAdGroupTimestamp(sheet, columnNumber) {
  let date = new Date();
  sheet.getRange(ROW_NUMBERS.LAST_RUN, columnNumber).setValue(date);
}

function getSettings(sheet) {
  let settings = {};
  for (let settingsKey in SETTING_CELL_REFERENCES) {
    const { reference, type } = SETTING_CELL_REFERENCES[settingsKey];
    const value = sheet.getRange(reference).getValue();
    const settingValue = getSettingValue(value, type);
    settings[settingsKey] = settingValue;
  }
  return settings;
}

function getSettingValue(value, settingType) {
  if (settingType === 'default') {
    return value;
  }
  if (settingType === 'csv') {
    return stringToCsv(value);
  }
  if (settingType === 'boolean') {
    return value;
  }
  throw new Error(`Setting type '${settingType}' not found`)
}

/*
* This is a customer request
* If the cell A18 is "Query Contains" then use the list below
* Otherwise just use cell H3 which is comma separated
* Returns an array
*/
function getQueryContains(sheet) {
  let useList = sheet.getRange("A18").getValue() === "Query Contains"
  if (!useList) return stringToCsv(sheet.getRange("H3").getValue())

  let list = []
  let rowNumber = 19
  while (sheet.getRange(rowNumber, 1).getValue()) {
    list.push(sheet.getRange(rowNumber, 1).getValue())
    rowNumber++
  }

  return list

}

/**
 * Clear ad group log
 * @param {SpreadsheetApp.sheet} sheet
 * @param {number} column
 */
function clearLog(sheet, column) {
  setColumnLog(sheet, column, "", "");
}


const SheetParams = {


  subRuleNumber: 1,

  /**
   * Get the sheet params
   * @param {Object} sheetParamsData 
   * @param {NegativeList} negativeList 
   * @returns {Object} sheetParams
   */
  getSheetParams(sheetParamsData, negativeList) {
    //for each ad group name or id (comma separated) add an extra rule
    //we'll create sub-rules from each rule (where there might be comma separated ad group names
    //or commands like RUN_ALL_ENABLED or RUN_FIRST_ENABLED)
    sheetParamsData = this._populateCampaignIds(sheetParamsData);
    let sheetParams = {};
    for (let ruleName in sheetParamsData) {
      sheetParams[ruleName] = sheetParams[ruleName] || {};
      if (!sheetParamsData[ruleName].campaignId) {
        continue;
      }
      sheetParams = this._createAdGroupSubRules(sheetParamsData[ruleName], sheetParams, ruleName,);
      sheetParams = this._createCampaignSubRules(sheetParamsData[ruleName], sheetParams, ruleName,);
      this._validateSheetParams(sheetParams[ruleName], Object.keys(negativeList.negativeListsMap));
    }
    return sheetParams;
  },

  /**
   * Validate the sheet params
   * @param {Object} sheetParams 
   * @param {Array} negativeListNames 
   */
  _validateSheetParams(ruleParams, negativeListNames) {
    for (let ruleName in ruleParams) {
      SubRuleValidationManager.validateSubRule(ruleParams[ruleName], negativeListNames);
    }
  },

  _populateCampaignIds(sheetParamsData) {
    for (let ruleName in sheetParamsData) {
      sheetParamsData[ruleName]['campaignId'] = sheetParamsData[ruleName]['campaignId'] || getCampaignIdFromName(sheetParamsData[ruleName].campaignName);
    }
    return sheetParamsData;
  },

  _createAdGroupSubRules(ruleParamsData, sheetParams, ruleName,) {
    let adGroupIds = getAdGroupIdsFromAdGroupString(ruleParamsData);
    testLog(`adGroupIds: ${JSON.stringify(adGroupIds)}`);
    for (let adGroupId of adGroupIds) {
      const key = `SUB_RULE_${this.subRuleNumber}`;
      sheetParams[ruleName][key] = JSON.parse(JSON.stringify(ruleParamsData));
      sheetParams[ruleName][key]['adGroupId'] = adGroupId;
      sheetParams[ruleName][key]['ruleName'] = ruleName;
      sheetParams[ruleName][key]['columnNumber'] = ruleParamsData.columnNumber;
      testLog(`ad group sub rule columnNumber: ${ruleParamsData.columnNumber}`);
      sheetParams[ruleName][key]['negativeListName'] = sheetParams[ruleName][key]['negativeListName'].split(",").map(x => x.trim());
      this.subRuleNumber++;
    }
    return sheetParams;
  },

  _createCampaignSubRules(ruleParamsData, sheetParams, ruleName) {
    if (ruleParamsData.adGroupId || ruleParamsData.adGroupName) {
      return sheetParams;
    }
    const key = `SUB_RULE_${this.subRuleNumber}`;
    sheetParams[ruleName][key] = JSON.parse(JSON.stringify(ruleParamsData));
    sheetParams[ruleName][key]['ruleName'] = ruleName;
    sheetParams[ruleName][key]['negativeListName'] = sheetParams[ruleName][key]['negativeListName'].split(",").map(x => x.trim());
    sheetParams[ruleName][key]['columnNumber'] = ruleParamsData.columnNumber;
    testLog(`campaign sub rule columnNumber: ${ruleParamsData.columnNumber}`);
    sheetParams[ruleName][key]['campaignId'] = sheetParams[ruleName][key]['campaignId'] || getCampaignIdFromName(ruleParamsData);
    this.subRuleNumber++;
    return sheetParams;
  },

}

const SubRuleValidationManager = {

  validateSubRule(subRule, negativeListNames) {
    this._validateNegativeListName(subRule, negativeListNames);
  },

  /**
   * Validate the negative list name
   * @param {Object} subRule 
   * @param {NegativeList} negativeList 
   */
  _validateNegativeListName(subRule, negativeListNames) {
    if (subRule.addNegativeTo === ADD_NEGATIVE_TO_OPTIONS.LIST && !subRule.negativeListName) {
      throw new Error(`Negative list name is required when add negative to is set to list (Row: ${ROW_NUMBERS.ADD_NEGATIVE_TO}, Rule: ${subRule.ruleName})`);
    }
    if (subRule.addNegativeTo === ADD_NEGATIVE_TO_OPTIONS.LIST && !negativeListNames.includes(subRule.negativeListName[0])) {
      throw new Error(`Negative list name '${subRule.negativeListName[0]}' not found in account. Lists found: ${negativeListNames.join(", ")}`);
    }
  }
}

function getCampaignIdFromName(campaignName) {
  if (!campaignName || typeof campaignName !== 'string' || campaignName.trim() === '') {
    return null;
  }
  if (new EntityFromPattern('CAMPAIGN', campaignName).isValidPattern()) {
    return new EntityFromPattern('CAMPAIGN', campaignName).getEntityId();
  }
  const query = `select campaign.id from campaign where campaign.name = '${campaignName}'`;
  const rows = AdsApp.report(query).rows();
  if (!rows.hasNext()) {
    throw new Error(`Campaign ${campaignName} not found`);
  }
  return rows.next()['campaign.id'];
}


const SheetParamsData = {
  /**
   * Get the rule params including:
   * adGroups, negativeLists, adGroupParams.numberOfMatches, adGroupColumnNumbers
   * @param {SpreadsheetApp.sheet} sheet 
   * @param {Object} runTimes
   * @returns {Object} ruleParams
   */
  getSheetParamsData(sheet, runTimes) {

    let sheetParamsData = {};

    for (let runTime of runTimes) {

      //grab adGroup data from sheet,store in arrays
      const columnNumber = runTime.columnNumber;

      sheetParamsData = this._getRuleParamsData(sheetParamsData, sheet, columnNumber);

    }
    return sheetParamsData;

  },

  _getRuleParamsData(sheetParamsData, sheet, columnNumber) {

    const ruleParamsData = SheetParamsData._readRuleData(sheet, columnNumber);
    const ruleName = `RULE_${columnNumber - 3}`;
    console.log(`ruleParamsData: ${JSON.stringify(ruleParamsData)}`);

    sheetParamsData[ruleName] = ruleParamsData;
    sheetParamsData[ruleName]['columnNumber'] = columnNumber;

    return sheetParamsData;
  },

  _readRuleData(sheet, columnNumber) {
    let ruleParamsData = {};
    for (let ruleParamKey in RULE_PARAM_KEYS) {
      // Get the row number from ROW_NUMBERS using the key from RULE_PARAM_KEYS
      const rowNumber = ROW_NUMBERS[RULE_PARAM_KEYS[ruleParamKey]];
      const value = sheet.getRange(rowNumber, columnNumber).getValue();
      // console.log(`ruleParamKey: ${ruleParamKey}, rowNumber: ${rowNumber}, value: ${value}`);
      // Initialize the array if it doesn't exist
      ruleParamsData[ruleParamKey] = value;
    }
    ruleParamsData.keywords = this._readKeywordData(sheet, columnNumber);
    return ruleParamsData;
  },
  _readKeywordData(sheet, columnNumber) {
    let keywords = []
    let row = ROW_NUMBERS.FIRST_POSITIVE_KEYWORD;
    while (sheet.getRange(row, columnNumber).getValue()) {
      keywords.push(sheet.getRange(row, columnNumber).getValue());
      row++;
    }
    return keywords;
  }
}

/**
 * If a comma separated list is provided, return those
 * Otherwise build up list names
 * @param {string} negativeKeywordListText 
 * @param {boolean} createListIfFull 
 * @returns {array} negative list names
 */
function updateNegativeKeywordListNames(negativeKeywordListText, createListIfFull) {

  if (negativeKeywordListText === "") {
    return [];
  }

  const negativeListsArray = negativeKeywordListText.split(",")
    .map(x => { return x.trim(); })
    .filter(x => { return x !== "" });

  if (!createListIfFull) {
    return [negativeListsArray[0]];
  }
  if (negativeListsArray.length === 0) {
    return [];
  }
  if (negativeListsArray.length > 1) {
    return negativeListsArray;
  }
  const firstListName = negativeListsArray[0];
  for (let i = 1; i < 10; i++) {
    let listName = `${firstListName} ${i}`;
    negativeListsArray.push(listName);
  }
  return negativeListsArray;

}

/**
 * Take the ad group string (value) from the sheet
 * convert it into multiple ad group ids if comma separated
 * or get all ad group ids from the campaign if RUN_ALL_ENABLED
 * @param {object} ruleParamsData 
 * @returns {array} adGroupIds
 */
function getAdGroupIdsFromAdGroupString(ruleParamsData) {
  if (ruleParamsData.adGroupName === AD_GROUP_COMMANDS.RUN_ALL_ENABLED) {
    return getAllAdGroupIds(ruleParamsData.campaignId);
  }
  if (ruleParamsData.adGroupName === AD_GROUP_COMMANDS.RUN_FIRST_ENABLED) {
    return [getAllAdGroupIds(ruleParamsData.campaignId, 1)[0]];
  }

  if (!ruleParamsData.adGroupName) {
    return [];
  }

  let adGroupNamesString = ruleParamsData.adGroupName.split(",").map(x => x.trim());

  let adGroupIds = [];
  for (let adGroupNameString of adGroupNamesString) {
    const adGroupId = new EntityFromPattern(ENTITY_TYPES.AD_GROUP, adGroupNameString, ruleParamsData.campaignId).getEntityId();
    adGroupIds.push(adGroupId);
  }
  return adGroupIds;
}
/**
 * Get all ad group ids from a given campaign name or id
 * @param {string} campaignId 
 * @returns {array} adGroupIds
 */
function getAllAdGroupIds(campaignId, limit = 0) {
  if (!campaignId) {
    throw new Error(`Campaign ID is required`);
  }
  let query = `select ad_group.id from ad_group where
			ad_group.status = 'ENABLED' and campaign.status = 'ENABLED'`;
  if (campaignId) {
    query += ` and campaign.id = '${campaignId}'`;
  }
  if (limit > 0) {
    query += ` limit ${limit}`;
  }
  const rows = AdsApp.report(query).rows();
  if (!rows.hasNext()) {
    return [];
  }
  let adGroupIds = [];
  while (rows.hasNext()) {
    let row = rows.next();
    adGroupIds.push(row["ad_group.id"]);
  }
  return adGroupIds;
}



/**
 * Get the keywords (the positive keywords) for 
 * @param {SpreadsheetApp.sheet} sheet
 * @param {number} column
 * @param {boolean} adGroupNamesAsKeywords
 * @param {string} adGroupName
 * @returns {Array} keywords
 */
function getAdGroupKeywords(sheet, column, adGroupNamesAsKeywords, adGroupName) {
  let keywords = [];

  if (adGroupNamesAsKeywords) {
    keywords = getKeywordsFromAdGroupName(adGroupName);
    log(
      "Note: words from Ad Group names will be as well as keywords in the settings sheet."
    );
  }

  const cellValues = getKeywordCellValues(sheet, column);

  //there may be multiple keywords in a single cell, comma separated
  const sheetKeywords = keywordCellValuesToKeywords(cellValues);

  keywords = keywords.concat(sheetKeywords);

  return keywords;
}

/**
 * Take the raw cell data, process into keywords
 * @param {array} cellValues 
 * @returns {array} keywords
 */
function keywordCellValuesToKeywords(cellValues) {
  let keywords = [];
  for (let cellValue of cellValues) {
    cellValue = String(cellValue).toLowerCase().trim();
    let cellKeywords = cellValue.split(",").map(x => {
      return x.trim();
    });
    for (let keyword of cellKeywords) {
      //skip duplicates
      if (keywords.indexOf(keyword) > -1) {
        continue;
      }
      keywords.push(keyword)
    }
  }
  return keywords;
}

/**
 * Get the raw cell values where the keywords are stored
 * @param {SpreadsheetApp.sheet} sheet 
 * @param {number} column 
 * @returns {array} keyword cell values
 */
function getKeywordCellValues(sheet, column) {
  let cellValues = [];
  let row = ROW_NUMBERS.FIRST_POSITIVE_KEYWORD;
  while (
    sheet
      .getRange(row, column)
      .getValue()
  ) {
    let cellValue = sheet.getRange(row, String(column)).getValue();
    cellValues.push(cellValue);
    row++;
  }
  return cellValues;
}

/**
 * Populate keywords from ad group name
 * @param {string} adGroupName 
 * @returns {Array} keywords
 */
function getKeywordsFromAdGroupName(adGroupName) {
  let keywords = [];

  keywords = adGroupName.split(" ").map(x => {
    return x.trim().toLowerCase();
  });
  log(
    "Note: words from Ad Group names will be as well as keywords in the settings sheet."
  );
  return keywords;
}

function runSheet(sheetName, negativeList, sheet) {
  console.log("Running sheet: " + sheetName);
  new SheetAdmin(sheet).update();


  let SETTINGS = getSettings(sheet);
  testLog("Settings: " + JSON.stringify(SETTINGS));

  const runTimes = new GetRunTimes(ROW_NUMBERS).getSheetLastRunTimes(sheet);
  testLog(`runTimes: ${JSON.stringify(runTimes)}`);
  const sheetParamsData = SheetParamsData.getSheetParamsData(sheet, runTimes);
  testLog(`sheetParamsData: ${JSON.stringify(sheetParamsData)}`);
  let sheetParams = SheetParams.getSheetParams(sheetParamsData, negativeList);
  testLog(`sheetParams: ${JSON.stringify(sheetParams)}`);

  const firstRuleCampaignId = sheetParamsData[Object.keys(sheetParamsData)[0]].campaignId;
  if (!sheetParamsData || shouldSkipSheet(firstRuleCampaignId,)) {
    return;
  }

  for (let ruleName in sheetParams) {
    if (Object.keys(sheetParams[ruleName]).length === 0) {
      continue;
    }
    try {
      runRule(sheetParams[ruleName], sheet, SETTINGS, negativeList);
    } catch (e) {
      console.error(`Error running rule ${ruleName}`, e);
      console.log(sheetParams[ruleName]);
      const columnNumber = sheetParams[ruleName]['SUB_RULE_1'].columnNumber || sheetParams[ruleName].columnNumber;
      testLog(`Adding error to column ${columnNumber}`);
      setColumnLog(sheet, columnNumber, "Error. Rule skipped.", e.message);
      continue;
    }
  }

  //timestamp
  let date = new Date();
  sheet.getRange(1, TIMESTAMP_COLUMN).setValue(date);


  log(`Finished running sheet: ${sheetName}`);
  log('---------------------------------------------\n');
}

function runRule(ruleParams, sheet, SETTINGS, negativeList) {
  let outputData = [];
  for (let subRuleName in ruleParams) {
    const subRuleParams = ruleParams[subRuleName];
    const negativesAdded = runSubRule(subRuleParams, sheet, SETTINGS, negativeList);
    console.log(`negativesAdded: ${JSON.stringify(negativesAdded)}`);

    if (negativesAdded.length === 0) {
      testLog(`No negatives added for adGroup: ${subRuleParams.adGroupName}`);
      continue;
    }
    testLog(`subRuleParams.logToSheetBool: ${subRuleParams.logToSheetBool}`);


    outputData.push({
      negativesAdded,
      adGroupName: subRuleParams.adGroupName,
      adGroupId: subRuleParams.adGroupId,
      numberOfMatches: subRuleParams.numberOfMatches,
      keywords: subRuleParams.keywords,
      campaignName: subRuleParams.campaignName,
    });
  }

  testLog('outputData: length: ' + outputData.length);
  testLog('outputData: ' + JSON.stringify(outputData));

  const firstRuleParams = ruleParams[Object.keys(ruleParams)[0]];
  if (firstRuleParams.logToSheetBool) {
    logNegativeToOutputSheet(outputData, firstRuleParams, sheet, SETTINGS.NEGATIVE_MATCH_TYPE);
  }

  if (outputData.length > 0) {
    new EmailAlert(sheet, SETTINGS, outputData, firstRuleParams).send();
  }
}

/**
 * Log the negatives added to the output sheet
 * @param {string} outputSheetUrl 
 * @param {array[]} negativesAdded 
 * @param {object} adGroupParams 
 * @param {string} rulesSheetName 
 * @param {string} matchType 
 */
function logNegativeToOutputSheet(outputData, ruleParams, rulesSheet, matchType) {
  const outputSpreadsheet = OutputSheet.getOutputSpreadsheet(ruleParams.outputSheetUrl, ruleParams, rulesSheet.getName());
  OutputSheet.setOutputSpreadsheetUrl(outputSpreadsheet.getUrl(), rulesSheet, ruleParams);
  const outputSheet = OutputSheet.getOutputSheet(outputSpreadsheet);
  outputSheet.clear();
  outputSheet.getRange('A1').setValue(new Date());
  outputSheet.getRange('A2').setValue('No negatives found');
  for (let subRuleIndex in outputData) {
    const subRuleOutputData = outputData[subRuleIndex];
    OutputSheet.setOutputValues(outputSheet, subRuleOutputData, ruleParams, parseInt(subRuleIndex) + 1, matchType);
  }
}

const OutputSheet = {


  getOutputSpreadsheet(outputSheetUrl, adGroupParams, rulesSheetName) {
    if (outputSheetUrl) {
      return SpreadsheetApp.openByUrl(outputSheetUrl)
    }
    console.log(`Creating output spreadsheet`)
    return this._createOutputSpreadsheet(adGroupParams, rulesSheetName);
  },

  _createOutputSpreadsheet(adGroupParams, rulesSheetName) {
    const name = `Auto Negative Keywords Ouput Sheet - ${AdsApp.currentAccount().getName()} - ${rulesSheetName} - Rule ${parseInt(adGroupParams.columnNumber) - 3}`;
    const spreadsheet = SpreadsheetApp.create(name);
    return spreadsheet;
  },

  getOutputSheet(outputSpreadsheet) {
    const outputSheet = outputSpreadsheet.getSheetByName(OUTPUT_SHEET_NAME);
    if (!outputSheet) {
      return outputSpreadsheet.getSheets()[0].setName(OUTPUT_SHEET_NAME);
    }
    return outputSpreadsheet.getSheetByName(OUTPUT_SHEET_NAME);
  },

  setOutputSpreadsheetUrl(outputSpreadsheetUrl, rulesSheet, adGroupParams) {
    rulesSheet.getRange(ROW_NUMBERS.OUTPUT_SHEET_URL, adGroupParams.columnNumber).setValue(outputSpreadsheetUrl);
    console.log(`Output sheet url: ${outputSpreadsheetUrl}`);
  },

  /**
   * 
   * @param {SpreadsheetApp.sheet} outputSheet 
   * @param {{
   *   negativesAdded: string[],
   *   adGroupName: string,
   *   numberOfMatches: number,
   *   keywords: string[],
   *   campaignName: string
   * }} subRuleOutputData
   * @param {number} columnNumber 
   */
  setOutputValues(outputSheet, subRuleOutputData, ruleParams, columnNumber, matchType) {
    testLog('Writing to output sheet');
    testLog('columnNumber: ' + columnNumber);
    const negativesLogArray = this._buildOutputValuesArray(subRuleOutputData);
    this._writeOutputValues(outputSheet, negativesLogArray, ruleParams, columnNumber, matchType);
  },

  _buildOutputValuesArray(subRuleOutputData,) {
    const negativesAdded = subRuleOutputData.negativesAdded.map(x => [x]);
    const adGroupName = subRuleOutputData.adGroupId ? getAdGroupNameFromId(subRuleOutputData.adGroupId) : undefined;
    let headers = [];
    if (adGroupName) {
      headers.push(['Ad Group: ' + adGroupName]);
      headers.push(['Ad Group ID: ' + subRuleOutputData.adGroupId]);
    }
    headers.push(['Negatives Added']);
    return [...headers, ...negativesAdded];
  },

  _writeOutputValues(outputSheet, negativesLogArray, ruleParams, columnNumber, matchType) {
    const startRow = 10;
    outputSheet.getRange(startRow, columnNumber, negativesLogArray.length, negativesLogArray[0].length).setValues(negativesLogArray);
    outputSheet.getRange('A1').setValue(new Date());
    outputSheet.getRange('A2').setValue(`Preview Mode: ${AdsApp.getExecutionInfo().isPreview()}`);
    outputSheet.getRange('A3').setValue(`Campaign Name: ${getCampaignNameFromId(ruleParams.campaignId)}`);
    outputSheet.getRange('A4').setValue(`Campaign ID: ${ruleParams.campaignId}`);
    outputSheet.getRange('A5').setValue(`Match type: ${matchType}`);
  }

}

function getAdGroupNameFromId(adGroupId) {
  const query = `select ad_group.name from ad_group where ad_group.id = '${adGroupId}'`;
  const rows = AdsApp.report(query).rows();
  if (!rows.hasNext()) {
    return '';
  }
  return rows.next()['ad_group.name'];
}

function getCampaignNameFromId(campaignId) {
  const query = `select campaign.name from campaign where campaign.id = '${campaignId}'`;
  const rows = AdsApp.report(query).rows();
  if (!rows.hasNext()) {
    return '';
  }
  return rows.next()['campaign.name'];
}


/**
 * 
 * @param {object} sheet 
 * @param {object} SETTINGS 
 * @param {array[]} negativesAdded 
 */

class EmailAlert {

  constructor(sheet, SETTINGS, outputData, firstRuleParams) {
    this.sheet = sheet;
    this.SETTINGS = SETTINGS;
    this.outputData = outputData;
    this.action = this.SETTINGS.PREVIEW_MODE || AdsApp.getExecutionInfo().isPreview() ? "found" : "added";
    this.campaignName = this.outputData[0].campaignName;
    this.firstRuleParams = firstRuleParams;
    this.timestampRange = this.sheet.getRange(1, EMAIL_TIMESTAMP_COLUMN);
  }

  send() {
    if (this.outputData.length === 0) {
      return;
    }
    if (!this.SETTINGS.EMAIL_ALERT) {
      return;
    }
    if (this._hasSentEmailToday()) {
      console.log("Email already sent today. Skipping.");
      return;
    }
    if (this.SETTINGS.EMAILS.length === 0) {
      console.warn("No emails provided in the settings. No email will be sent.");
      return;
    }
    const html = this._getHtml();
    const subject = this._getSubject();
    this._sendEmail(html, subject);
    this._logTimeStamp();
  }

  _getSubject() {
    const accountName = AdsApp.currentAccount().getName();
    //get the number of negatives added in total
    const totalNegativesAdded = this.outputData.reduce((acc, val) => acc + val.negativesAdded.length, 0);
    return `${totalNegativesAdded} negatives were ${this.action} - ${this.sheet.getName()} sheet - ${accountName}`
  }

  _getHtml() {
    const numberOfAdGroups = this.outputData.length;
    let html = "<p>Hello,</p>";
    html += "<p></br>The Auto Negative Keywords Pro Script ran successfully</p>";
    html += `<p></br>Settings sheet url: ${INPUT_SHEET_URL}</p>`;
    html += `<p><b>Sheet name:</b> ${this.sheet.getName()}</p>`;
    html += `<br/>`;
    html += `<p><b>Notes:</b></p>`;

    if (AdsApp.getExecutionInfo().isPreview()) {
      html += `<p>The script was previewed. No changes were made.</p>`;
    }
    if (this.SETTINGS.PREVIEW_MODE && !AdsApp.getExecutionInfo().isPreview()) {
      html += `<p>Preview Mode was enabled in the settings. No changes were made.</p>`;
    }
    if (!this.SETTINGS.PREVIEW_MODE && !AdsApp.getExecutionInfo().isPreview()) {
      html += `<p>The negative keywords were successfully added to the account.</p>`;
    }
    html += `<p>Only one email per sheet per day will be sent.</p>`;
    html += `<br/>`;

    html += `<p>Here are the relevant rules:</p>`;

    for (let outputData of this.outputData) {
      html += `<br/>`;
      html += `<p><b>First 50 negatives:</b> ${outputData.negativesAdded.slice(0, 50).join(", ")}</p>`;
      if (outputData.negativesAdded.length > 50) {
        html += `<p><b>View the output sheets to view the rest.</b></p>`;
      }
      html += `<br/>`;
    }
    html += `<br/>`;
    html += `<p>For support visit the <a href="https://shabba.io/script/1"> script page</a> or respond to this email!</p>`
    return html;
  }

  _sendEmail(html, subject) {
    for (let email of this.SETTINGS.EMAILS) {
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: html,
        replyTo: "charles@shabba.io"
      });
    }
  }

  _logTimeStamp() {
    const date = new Date();
    this.timestampRange.setValue(date);
  }

  _getTimestamp() {
    return this.timestampRange.getValue();
  }

  _hasSentEmailToday() {
    const timestamp = new Date(this._getTimestamp());
    if (!timestamp) {
      return false;
    }
    const today = new Date();
    return today.toDateString() === timestamp.toDateString();
  }

}

/**
 * Whether to skip a sheet based on the first rule's campaign
 * @param {string} campaignId 
 * @returns 
 */
function shouldSkipSheet(campaignId,) {
  const campaignInfo = new CampaignInfo();
  if (campaignInfo.getCampaignStatusFromId(campaignId) === STATUSES.PAUSED) {
    console.warn(`Campaign '${campaignInfo.campaignName}' is paused. Skipping.`);
    return true;
  }

  if (campaignInfo.getCampaignStatusFromId(campaignId) === STATUSES.NOT_FOUND) {
    console.warn(`Campaign '${campaignInfo.campaignName}' could not be found. Skipping.`);
    return true;
  }
  return false;
}

function runSubRule(subRuleParams, sheet, SETTINGS, negativeList) {
  testLog(`Running ad group: ${JSON.stringify(subRuleParams)}`);
  testLog(`${subRuleParams.keywords.length} positive keywords found`);
  clearLog(sheet, subRuleParams.columnNumber);
  getAdGroupNegatives(subRuleParams, SETTINGS, sheet);
  testLog(`subRuleParams: ${JSON.stringify(subRuleParams)}`);
  const negativesAdded = processNegativeKeywords(subRuleParams, SETTINGS, sheet, negativeList);
  setAdGroupTimestamp(sheet, subRuleParams.columnNumber);
  return negativesAdded;
}

function skipAGroup(numberOfMatches, ruleParams, adGroupId, sheet, column, SETTINGS) {

  if (numberOfMatches == "") {
    let message =
      "Number of matches not set for the Ad Group '" +
      adGroupId +
      "'. The Ad Group will be skipped.";
    log(message);
    sendErrorEmail(message, SETTINGS["EMAILS"]);
    setColumnLog(sheet, column, "Error. Ad Group skipped.", message);
    return true;
  }

  testLog("'Positive keywords' from sheet: " + ruleParams.keywords);
  if (!SETTINGS['RUN_KEYWORDLESS_AD_GROUPS'] && ruleParams.keywords.length == 0) {
    setColumnLog(
      sheet,
      column,
      "Please add ruleParams.keywords below (no negative ruleParams.keywords were added)",
      "",
      true
    );
    return true;
  }
  return false;
}

/**
 * Get the initial negative keywords based on the settings
 * we'll filter further later
 * @param {Object} ruleParams
 * @param {Object} SETTINGS
 * @param {Object} sheet
 * 
 */
function addAdGroupNegativesToParams(ruleParams, SETTINGS, sheet) {
  const adGroupId = ruleParams.adGroupId;
  const columnNumber = ruleParams.columnNumber;
  const numberOfMatches = ruleParams.numberOfMatches;
  const keywords = ruleParams.keywords;
  const campaignId = ruleParams.campaignId;

  if (skipAGroup(numberOfMatches, ruleParams, adGroupId, sheet, columnNumber, SETTINGS)) {
    return
  }

  const negativeKeywords = generateNegativesFromAccount(
    SETTINGS,
    keywords,
    numberOfMatches,
    adGroupId,
    campaignId, //get the list from the first campaign in the list,
    ruleParams.approxMatch,
    ruleParams.approxMatchThreshold
  );

  ruleParams.negativeKeywords = negativeKeywords;
}

/**
 * Add a negativeKeywords list to subRuleParams
 * @param {Object} subRuleParams
 * @param {Object} SETTINGS
 * @param {Object} sheet
 * @param {Object} negativeList
 */
function getAdGroupNegatives(subRuleParams, SETTINGS, sheet) {
  //loop through the adGroups listed in the sheet

  addAdGroupNegativesToParams(subRuleParams, SETTINGS, sheet)
  testLog(`Found ${subRuleParams.negativeKeywords.length} negative keywords pre-filtering`)
  testLog(`subRuleParams.negativeKeywords: ${JSON.stringify(subRuleParams.negativeKeywords)}`);
  filterUnnecessaryNegativeKeywords(subRuleParams, SETTINGS)
  testLog(`Found ${subRuleParams.negativeKeywords.length} negative keywords post-filtering`)

}

/**
 * Filters out negative keywords which won't have an affect
 * e.g. it's already in as a negative keyword
 * @param {object} subRuleParams 
 * @param {object} SETTINGS 
 */
function filterUnnecessaryNegativeKeywords(subRuleParams, SETTINGS) {
  const negativeKeywords = subRuleParams.negativeKeywords;
  if (negativeKeywords.length == 0) {
    return;
  }

  const adGroupId = subRuleParams.adGroupId;
  const campaignId = subRuleParams.campaignId;

  let existingNegatives = new ExistingNegativeKeywords(
    campaignId,
    adGroupId,
    subRuleParams.addNegativeTo,
    SETTINGS
  ).get();

  testLog(`existingNegatives: ${JSON.stringify(existingNegatives)}`);

  subRuleParams.negativeKeywords = subRuleParams.negativeKeywords.filter(existingNegative => existingNegative !== '');
  for (let negativeKeyword of negativeKeywords) {
    if (!keywordIsBlocked(negativeKeyword, SETTINGS['NEGATIVE_MATCH_TYPE'], existingNegatives)) {
      continue;
    }
    console.log(`Negative keyword '${negativeKeyword}' is blocked. Removing from subRuleParams.negativeKeywords`);
    subRuleParams.negativeKeywords = subRuleParams.negativeKeywords.filter(existingNegative => existingNegative !== negativeKeyword);
  }


}



function processNegativeKeywords(subRuleParams, SETTINGS, sheet, negativeList) {

  let negativesAdded = [];

  addAdGroupNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, subRuleParams.campaignId);
  addCampaignNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, subRuleParams.campaignId);
  processListNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, negativeList);
  updateLogsWithNegativesAdded(negativesAdded, sheet, subRuleParams, SETTINGS['PREVIEW_MODE']);
  return negativesAdded;

}


function processListNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, negativeList, negativeListsIndex = 0) {
  if (subRuleParams.addNegativeTo !== ADD_NEGATIVE_TO_OPTIONS.LIST) {
    testLog(`Skipping adding negatives to list as addNegativeTo is set to ${subRuleParams.addNegativeTo}`)
    return;
  }
  if (subRuleParams.negativeListName[negativeListsIndex] === "") {
    return;
  }
  if (!subRuleParams.negativeListName[negativeListsIndex]) {
    return;
  }
  if (negativeListsIndex >= subRuleParams.negativeListName.length) {
    return;
  }

  let negativeListName = subRuleParams.negativeListName[negativeListsIndex];
  if (!SETTINGS["CREATE_LIST_IF_FULL"] && negativeList.negativeListsMap[negativeListName].isFull) {
    console.warn(`Negative list '${negativeListName}' is full. Skipping.`);
    return;
  }

  //use the list unless full then move onto the next one
  //create if not exists
  negativeList.createNegativeListIfNotExists(negativeListName, subRuleParams.negativeListName[0]);

  if (negativeList.negativeListsMap[negativeListName].isFull) {
    negativeListsIndex++;
    log(`Negative list '${negativeListName}' is full. Moving onto the next list (${subRuleParams.negativeListName[negativeListsIndex]})`);
    processListNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, negativeList, negativeListsIndex);
    return;
  }
  negativeList.addNegativeKeywordsFactory(
    negativeListName,
    subRuleParams.negativeKeywords,
    negativesAdded,
    SETTINGS['NEGATIVE_MATCH_TYPE'],
    SETTINGS['PREVIEW_MODE'],
  )
}
class AddListNegatives {

  constructor() {
    this.negativesAddedCount = 0;
    this.remainingSpace = 0;
    this.negativesToAddCount = 0;
    this.maxListLength = 5000;
  }

  add(params, negativesAdded) {

    const { negativeKeywords, matchType, isPreview, list } = params;

    let negativesWithMatchType = this.addMatchTypeToKeywords(negativeKeywords, matchType);
    this.setRemainingSpace(list.negatives);
    this.setNegativesToAddCount(negativeKeywords);
    let negativesChunk = this.getNegativesChunk(negativesWithMatchType);
    for (let negative of negativesChunk) {
      negativesAdded.push(negative);
    }
    this.addNegativeKeywords(isPreview, list['list'], negativesChunk);
    testLog(`${negativesChunk.length} negatives were added`);
    this.negativesAddedCount += negativesChunk.length;
    if (this.negativesToAddCount > this.negativesAddedCount) {
      testLog(`Total negatives to add: ${this.negativesToAddCount}. Negs added: ${this.negativesAddedCount}. Running again.`);
      const nextNegativeKeywordsBatch = negativeKeywords.slice(negativesChunk.length);
      testLog(`Next batch length: ${nextNegativeKeywordsBatch.length}`)
      params.negativeKeywords = nextNegativeKeywordsBatch;
      this.add(params, negativesAdded);
    }
  }

  addNegativeKeywords(isPreview, negativeKeywordList, negativeKeywords) {
    const inRunMode = !AdsApp.getExecutionInfo().isPreview();
    if (isPreview && inRunMode) {
      return;
    }
    negativeKeywordList.addNegativeKeywords(negativeKeywords);
  }

  setNegativesToAddCount(negatives) {
    if (this.negativesToAddCount > 0) {
      return;
    }
    this.negativesToAddCount = negatives.length;
  }

  setRemainingSpace(listNegatives) {
    if (this.remainingSpace > 0) {
      return;
    }
    this.remainingSpace = this.maxListLength - listNegatives.length;
  }

  getNegativesChunk(negatives) {
    let negativesChunk = negatives.slice(0, this.remainingSpace);
    return negativesChunk;
  }

  addMatchTypeToKeywords(negativeKeywordsChunk, matchType) {
    let negativesWithMatchType = [];
    for (let negativeKeywordIndex in negativeKeywordsChunk) {
      const negativeKeyword = negativeKeywordsChunk[negativeKeywordIndex];
      const negativeKeywordWithMatchType = addMatchType(
        negativeKeyword,
        matchType
      );
      negativesWithMatchType.push(negativeKeywordWithMatchType);
    }
    return negativesWithMatchType;
  }

}

class NegativeList {

  constructor() {
    this.negativeListsMap = {};
    this.maxListSize = 5000;
    this.negativeKeywordsAddedCount = [];
    this.remainingSpace = 0;
    this.totalNegativesToAddCount = 0;
  }

  /**
   * Add negative keywords to a negative list
   * @param {string} negativeListName 
   * @param {array} negativeKeywords 
   * @param {array} negativesAdded 
   * @param {string} matchType
   * @param {boolean} isPreview //from the sheet
   */
  addNegativeKeywordsFactory(negativeListName, negativeKeywords, negativesAdded, matchType, isPreview) {
    testLog(`negativeListName: ${negativeListName}`)
    testLog(`negativeKeywords: ${negativeKeywords}`)
    if (!negativeKeywords.length) {
      return;
    }
    let params = {
      negativeKeywords,
      matchType,
      isPreview,
      list: this.negativeListsMap[negativeListName]
    }
    new AddListNegatives().add(params, negativesAdded);
  }

  /*
  * Creates a new negative keyword list
  * if it doesn't already exist
  * @param {string} negativeListName
  * @param {string} sourceNegativeListName - where to grab the campaigns from
  */
  createNegativeListIfNotExists(negativeListName, sourceNegativeListName) {
    if (this.negativeListsMap[negativeListName]) {
      return;
    }

    testLog(`Creating new keyword list '${negativeListName}'`)
    testLog(`${negativeListName} is not in ${Object.keys(this.negativeListsMap)}`)
    testLog(Object.keys(this.negativeListsMap).includes(negativeListName))
    let negativeKeywordListOperation = AdsApp.newNegativeKeywordListBuilder()
      .withName(negativeListName)
      .build();
    let negativeKeywordList = negativeKeywordListOperation.getResult();
    //assign the campaigns from the source list
    //source = first negative list in the comma separated list
    this.assignListToCampaigns(negativeKeywordList, this.negativeListsMap[sourceNegativeListName]["campaigns"]);
    this.negativeListsMap[negativeListName] = { list: negativeKeywordList, negatives: [], isFull: false };
  }

  /* 
  * Creates an object containing negative keyword list functions
  * Along with keywords therein
  * This prevents needing to create an iterator for each negativeKeyword added
  */
  createNegativeListsMap() {
    this.negativeListsMap = {};
    let listIter = AdsApp.negativeKeywordLists().get();
    while (listIter.hasNext()) {
      let negativeList = listIter.next();
      let negativeListName = negativeList.getName();
      this.negativeListsMap[negativeListName] = {};
      this.negativeListsMap[negativeListName]["list"] = negativeList;
      let sharedNegativeKeywordIterator = negativeList
        .negativeKeywords()
        .get();

      let sharedNegativeKeywords = [];

      while (sharedNegativeKeywordIterator.hasNext()) {
        let sharedNegativeKeyword = sharedNegativeKeywordIterator.next();
        sharedNegativeKeywords.push(sharedNegativeKeyword.getText());
      }

      this.negativeListsMap[negativeListName]["negatives"] = sharedNegativeKeywords;
      this.negativeListsMap[negativeListName]["isFull"] = sharedNegativeKeywords.length >= this.maxListSize;
      this.negativeListsMap[negativeListName]["campaigns"] = negativeList.campaigns().get();
    }
  }

  /**
   * Assign a negative keyword list to campaigns
   * @param {AdsApp.​NegativeKeywordList} negativeKeywordList 
   * @param {AdsApp.Campaigns} campaigns 
   */
  assignListToCampaigns(negativeKeywordList, campaigns) {
    while (campaigns.hasNext()) {
      let campaign = campaigns.next();
      campaign.addNegativeKeywordList(negativeKeywordList);
    }
  }

}


function updateLogsWithNegativesAdded(negativesAdded, sheet, ruleParams, previewMode) {
  // console.log(`updateLogsWithNegativesAdded > negativesAdded: ${JSON.stringify(negativesAdded)}`);
  let note = AdsApp.getExecutionInfo().isPreview() || previewMode
    ? "Ran in preview mode\n"
    : "Ran in production (not preview mode)\n";
  note +=
    negativesAdded.length > 0
      ? "Examples (1,000 max): \n" +
      negativesAdded.slice(0, 1000).join("\n")
      : "No new negative keywords were added";
  let logText = negativesAdded.length + " new negatives added."
  setColumnLog(
    sheet,
    ruleParams.columnNumber,
    logText,
    note
  );
}

function addAdGroupNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, campaignId, campaignType = campaignTypes.search) {

  if (subRuleParams.addNegativeTo !== ADD_NEGATIVE_TO_OPTIONS.AD_GROUP) {
    testLog(`Skipping adding negatives to ad group as addNegativeTo is set to ${subRuleParams.addNegativeTo}`)
    return;
  }
  if (subRuleParams.addNegativeTo === ADD_NEGATIVE_TO_OPTIONS.AD_GROUP && !subRuleParams.adGroupId) {
    throw new Error(`'Add Negative To' is 'Ad Group' yet no ad group ID found (Rule ${subRuleParams.ruleName})`)
  }

  let adGroupIterator;
  if (campaignType === campaignTypes.shopping) {
    adGroupIterator = AdsApp.shoppingAdGroups();
  } else {
    adGroupIterator = AdsApp.adGroups();
  }

  testLog(`subRuleParams.adGroupId: ${subRuleParams.adGroupId}`)
  testLog(`subRuleParams: ${JSON.stringify(subRuleParams)}`)

  adGroupIterator = adGroupIterator
    .withIds([subRuleParams.adGroupId])
    .withCondition("CampaignId = '" + campaignId + "'")
    .get();

  if (!adGroupIterator.hasNext() && campaignType === campaignTypes.search) {
    //this will happen if it's a shopping campaign
    addAdGroupNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, campaignId, campaignTypes.shopping);
    return;
  }

  if (!adGroupIterator.hasNext() && campaignType === campaignTypes.shopping) {
    console.error(`No ad group found for '${getAdGroupNameFromId(subRuleParams.adGroupId)}' (${subRuleParams.adGroupId}) in campaign '${subRuleParams.campaignName}' (${campaignId})`)
    return;
  }

  while (adGroupIterator.hasNext()) {
    let adGroup = adGroupIterator.next();
    adGroupNegativeKeywordService(adGroup, subRuleParams.negativeKeywords, SETTINGS, negativesAdded);
  }

}

function addCampaignNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, campaignId, campaignType = campaignTypes.search) {
  if (subRuleParams.addNegativeTo !== ADD_NEGATIVE_TO_OPTIONS.CAMPAIGN) {
    testLog(`Skipping adding negatives to campaign as addNegativeTo is set to ${subRuleParams.addNegativeTo}`)
    return;
  }

  let campaignIterator = campaignType === campaignTypes.shopping ? AdsApp.shoppingCampaigns() : AdsApp.campaigns();

  campaignIterator = campaignIterator
    .withIds([campaignId])
    .get();

  if (!campaignIterator.hasNext() && campaignType === campaignTypes.search) {
    //this will happen if it's a shopping campaign
    addCampaignNegativeKeywords(subRuleParams, SETTINGS, negativesAdded, campaignId, campaignTypes.shopping);
    return;
  }

  if (!campaignIterator.hasNext() && campaignType === campaignTypes.shopping) {
    console.error(`No campaign found for '${subRuleParams.campaignName}' (${campaignId})`)
    return;
  }

  while (campaignIterator.hasNext()) {
    let campaign = campaignIterator.next();
    campaignNegativeKeywordService(campaign, subRuleParams.negativeKeywords, SETTINGS, negativesAdded);
  }
}

/**
 * 
 * @param {AdsApp.AdGroup} adGroup 
 * @param {Array} negativeKeywords 
 * @param {Object} SETTINGS 
 * @param {Array} negativesAdded 
 */
function adGroupNegativeKeywordService(adGroup, negativeKeywords, SETTINGS, negativesAdded) {
  for (let negativeKeyword in negativeKeywords) {
    negativeKeyword = addMatchType(
      negativeKeywords[negativeKeyword],
      SETTINGS["NEGATIVE_MATCH_TYPE"]
    );
    if (!SETTINGS.PREVIEW_MODE || AdsApp.getExecutionInfo().isPreview()) {
      adGroup.createNegativeKeyword(negativeKeyword);
    }
    if (!negativesAdded.includes(negativeKeyword)) {
      negativesAdded.push(negativeKeyword);
    }
  }
}

function campaignNegativeKeywordService(campaign, negativeKeywords, SETTINGS, negativesAdded) {
  for (let negativeKeyword in negativeKeywords) {
    negativeKeyword = addMatchType(
      negativeKeywords[negativeKeyword],
      SETTINGS["NEGATIVE_MATCH_TYPE"]
    );
    if (!SETTINGS.PREVIEW_MODE || AdsApp.getExecutionInfo().isPreview()) {
      campaign.createNegativeKeyword(negativeKeyword);
    }
    if (!negativesAdded.includes(negativeKeyword)) {
      negativesAdded.push(negativeKeyword);
    }
  }
}

/*
 * Get all negative keywords from an AdGroup
 * Including AdGroup level, Campaign level and Negative Keyword List (account) level
 */
class ExistingNegativeKeywords {


  constructor(campaignId, adGroupId, addNegativeToOption, SETTINGS) {
    this.campaignId = campaignId;
    this.adGroupId = adGroupId;
    this.addNegativeToOption = addNegativeToOption;
    this.SETTINGS = SETTINGS;
    this.existingNegatives = [];
  }


  get() {

    if (!this.SETTINGS["NEW_NEGATIVES_ONLY"]) {
      return [];
    }

    if (this.addNegativeToOption === ADD_NEGATIVE_TO_OPTIONS.AD_GROUP) {
      this.getAdGroupNegatives();
      this.getCampaignNegatives();
    }
    if (this.addNegativeToOption === ADD_NEGATIVE_TO_OPTIONS.CAMPAIGN) {
      this.getCampaignNegatives();
    }
    this.getNegativeListNegatives();

    return this.existingNegatives;

  }

  getAdGroupNegatives() {
    const query = this.getAdGroupNegativesReportQuery();
    console.log("ad group negative keywords query: " + query);
    return this.getReportNegatives(query);
  }

  getCampaignNegatives() {
    const query = this.getCampaignNegativesReportQuery();
    return this.getReportNegatives(query);
  }

  getAdGroupNegativesReportQuery() {
    let report_name = "KEYWORDS_PERFORMANCE_REPORT";
    let query =
      "SELECT KeywordMatchType,Criteria " + " FROM " + report_name;

    query += " where IsNegative = TRUE ";
    query += " and CampaignId = '" + this.campaignId + "'";
    query += " and AdGroupId = '" + this.adGroupId + "'";
    return query;
  }

  getReportNegatives(query) {
    let report = AdsApp.report(query);

    let rows = report.rows();
    //loop through this campaign's queries, add anything which doesn't contain our positive keywords to the negativeKeywords array (these will be added as negatives later)
    while (rows.hasNext()) {
      let row = rows.next();

      let negative = {
        keyword: row.Criteria,
        match_type: row.KeywordMatchType.toLowerCase()
      };

      this.existingNegatives.push(negative);
    }
  }

  getCampaignNegativesReportQuery() {
    const report_name = "CAMPAIGN_NEGATIVE_KEYWORDS_PERFORMANCE_REPORT";
    let query =
      "SELECT KeywordMatchType,Criteria " + " FROM " + report_name;

    query += " where IsNegative = TRUE ";
    query += " and CampaignId = '" + this.campaignId + "'";

    return query;;
  }

  getNegativeListNegatives() {
    let campaigns = AdsApp.shoppingCampaigns()
      .withCondition('CampaignId = "' + this.campaignId + '"')
      .get();
    if (!campaigns.hasNext()) {
      campaigns = AdsApp.campaigns()
        .withCondition('CampaignId = "' + this.campaignId + '"')
        .get();
    }
    if (!campaigns.hasNext()) {
      console.log("Campaign name not found when looking for negative keyword lists")
      return
    };
    let campaign = campaigns.next();
    let listIter = campaign.negativeKeywordLists().get();
    if (!listIter.hasNext()) {
      console.log(`No negative lists were found for the campaign ${this.campaignName}`)
    }
    while (listIter.hasNext()) {
      let negativeList = listIter.next();
      let sharedNegativeKeywordIterator = negativeList
        .negativeKeywords()
        .get();
      while (sharedNegativeKeywordIterator.hasNext()) {
        let sharedNegativeKeyword = sharedNegativeKeywordIterator.next();
        this.existingNegatives.push({
          keyword: sharedNegativeKeyword.getText(),
          match_type: sharedNegativeKeyword
            .getMatchType()
            .toLowerCase()
        });
      }
    }
  }

}



function sendErrorEmail(msg, emails, body = "") {
  for (let email_i in emails) {
    let email = emails[email_i];
    MailApp.sendEmail({
      to: email,
      subject: msg,
      htmlBody: `${body}\n\nControl sheet url: ${INPUT_SHEET_URL}`
    });
  }
}

/**
 * Pull search terms and add them to the negatives list if necessary
 * @param {Object} SETTINGS
 * @param {Array} keywords
 * @param {Array} numberOfMatches
 * @param {string} adGroupId
 * @param {string} campaignId
 * @returns {Array} negativeKeywords
 */
function generateNegativesFromAccount(
  SETTINGS,
  keywords,
  numberOfMatches,
  adGroupId,
  campaignId,
  shouldUseFuzzyMatchScore,
  fuzzyMatchThreshold
) {

  const query = getGAQLQuery(SETTINGS, campaignId, adGroupId);

  log("query: " + query);
  let report = AdsApp.report(query);
  // AdsApp.report(query).exportToSheet(testOutputSheet);
  let rows = report.rows();
  console.log(`The search term report query returned ${rows.totalNumEntities()} rows`);
  let negativeKeywords = [];
  //loop through this campaign's queries, add anything which doesn't contain our positive keywords to the negativeKeywords array (these will be added as negatives later)
  let searchTerms = [];
  while (rows.hasNext()) {
    let row = rows.next();
    let query = row.Query;
    searchTerms.push(query);
    const validator = new SearchTermValidator(query, parseInt(SETTINGS["MIN_WORDS"]), parseInt(SETTINGS["MAX_WORDS"]));
    const isValid = validator.isValid();
    if (!isValid.valid) {
      if (LOG_INVALID_SEARCH_TERMS) {
        console.log(`Search term '${query}' is invalid and will be skipped: ${isValid.reason}`);
      }
      continue;
    }

    if (
      SETTINGS["RUN_KEYWORDLESS_AD_GROUPS"] &&
      keywords.length == 0
    ) {
      negativeKeywords.push(query);
      continue;
    }
    //console log what we're checking

    if (new SearchTermChecker(query, keywords, numberOfMatches, fuzzyMatchScore, shouldUseFuzzyMatchScore, fuzzyMatchThreshold,).isNegative()) {
      negativeKeywords.push(query);
    }
  }
  testLog(`Search Terms: ${searchTerms.join(", ")}`);
  console.log(`${negativeKeywords.length} negative keywords were found.`)
  return negativeKeywords;
}


class SearchTermValidator {

  /**
   * Invalid search terms will be skipped
   * We'll also return a reason for invalidity
   * @param {*} searchTerm 
   * @param {*} minWords 
   * @param {*} maxWords
   */

  constructor(query, minWords, maxWords) {
    this.query = query;
    this.minWords = minWords;
    this.maxWords = maxWords || 10;
  }

  /**
   * Check if the search term is valid
   * @returns {boolean}
   */
  isValid() {
    const words = this.query.split(" ");
    if (words.length < this.minWords) {
      return { valid: false, reason: "Search term is too short" };
    }
    if (words.length > this.maxWords) {
      return { valid: false, reason: "Search term is too long" };
    }
    return { valid: true, reason: "" };
  }
}

const PULL_FROM_OPTIONS = {
  'AD_GROUP_CAMPAIGN': 'AdGroup/Campaign',
  'ACCOUNT': 'Account',
}

/**
 * Get the GAQL query for the search terms
 * @param {Object} SETTINGS
 * @param {string} campaignId
 * @param {string} adGroupId
 * @returns {string} query
 */
function getGAQLQuery(SETTINGS, campaignId, adGroupId) {
  let query =
    "SELECT Query " +
    " FROM SEARCH_QUERY_PERFORMANCE_REPORT where Clicks < 100000 ";

  if (
    SETTINGS["PULL_FROM"] === PULL_FROM_OPTIONS.AD_GROUP_CAMPAIGN
  ) {
    query += ' AND CampaignId = "' + campaignId + '"';
    if (adGroupId) {
      query += ' AND AdGroupId = "' + adGroupId + '"';
    }
  }


  if (String(SETTINGS["CONVERSIONS_LESS_THAN"]) !== "") {
    query += " AND Conversions < " + SETTINGS["CONVERSIONS_LESS_THAN"];
  }
  if (String(SETTINGS["CLICKS_MORE_THAN"]) !== "") {
    query += " AND Clicks > " + SETTINGS["CLICKS_MORE_THAN"];
  }
  if (String(SETTINGS["IMPRESSIONS_MORE_THAN"]) !== "") {
    query += " AND Impressions > " + SETTINGS["IMPRESSIONS_MORE_THAN"];
  }

  for (let query_contains_index in SETTINGS["CONTAINS"]) {
    let contains_query = SETTINGS["CONTAINS"][query_contains_index];
    query += " and Query CONTAINS_IGNORE_CASE '" + contains_query + "'";
  }
  for (let query_not_contains_index in SETTINGS["NOT_CONTAINS"]) {
    let not_contains_query =
      SETTINGS["NOT_CONTAINS"][query_not_contains_index];
    query +=
      " and Query DOES_NOT_CONTAIN_IGNORE_CASE '" +
      not_contains_query +
      "'";
  }

  if (typeof SETTINGS["DATE_RANGE"] !== "number") {
    throw new Error("Lookback window (Days) is required");
  }
  query += ` DURING ${getAdWordsFormattedDate(SETTINGS["DATE_RANGE"], "yyyyMMdd")}, ${getAdWordsFormattedDate(0, "yyyyMMdd")}`;
  return query;
}

class SearchTermChecker {

  constructor(searchTerm, keywords, numberOfMatches, fuzzyMatchScoreFn, shouldUseFuzzyMatchScore, fuzzyMatchThreshold) {
    this.searchTerm = searchTerm;
    this.keywords = keywords;
    this.numberOfMatches = numberOfMatches;
    this.fuzzyMatchScoreFn = fuzzyMatchScoreFn;
    this.shouldUseFuzzyMatchScore = shouldUseFuzzyMatchScore;
    this.fuzzyMatchThreshold = fuzzyMatchThreshold;
  }

  /**
   * Check if a search term should be a negative
   * @returns {boolean}
   */
  isNegative() {
    let matches = 0;
    //loop through the positive keywords (from the sheet)
    for (let keywordIndex in this.keywords) {

      //the keyword might be comma separated so split it and check all of them
      let positiveKeywords = this.keywords[keywordIndex].split(",").map(x => {
        return x.trim().toLowerCase();
      });
      for (let positiveKeywordsIndex in positiveKeywords) {
        let positiveKeyword = positiveKeywords[positiveKeywordsIndex];

        if (this._hasMatch(this.searchTerm, positiveKeyword)) {
          matches++;
        }

      }
    }
    if (matches < this.numberOfMatches) {
      return true;
    }
    return false;
  }


  /*
   * Consider a match if the word contains the word e.g. 'shoes' is in 'red shoes'
   * Also treat the searchTerm as a single word e.g. check if 'red shoes' is in 'redshoes'
   * Then treat the positive keyword as a single word e.g. check if 'redshoes' is in 'red shoes'
   */
  _hasMatch(searchTerm, positiveKeyword) {

    if (this._isExactMatchPositiveWord(positiveKeyword)) {
      return this._checkExactMatch(positiveKeyword, searchTerm)
    }

    if (positiveKeyword.indexOf("regex(") > -1) {
      return this._checkRegexMatch(positiveKeyword, searchTerm);
    }

    if (searchTerm.indexOf(positiveKeyword) > -1) {
      return true;
    }

    if (this.shouldUseFuzzyMatchScore) {
      const thresholdFloat = parseFloat(this.fuzzyMatchThreshold) / 100;
      const fuzzyMatchScore = this.fuzzyMatchScoreFn(positiveKeyword, searchTerm);
      if (fuzzyMatchScore > thresholdFloat) {
        return true;
      }
      const spacelessFuzzyMatchScore = this.fuzzyMatchScoreFn(positiveKeyword, searchTerm.replace(" ", ""));
      if (spacelessFuzzyMatchScore > thresholdFloat) {
        return true;
      }
      for (let searchTermWord of searchTerm.split(" ")) {
        const fuzzyMatchScore = this.fuzzyMatchScoreFn(positiveKeyword, searchTermWord);
        if (fuzzyMatchScore > thresholdFloat) {
          return true;
        }
      }
      let startPosition = 0;
      let endPosition = positiveKeyword.length;
      while (endPosition < searchTerm.length) {
        const fuzzyMatchScore = this.fuzzyMatchScoreFn(positiveKeyword, searchTerm.substring(startPosition, endPosition));
        if (fuzzyMatchScore > thresholdFloat) {
          return true;
        }
        startPosition++;
        endPosition++;
      }
    }

    if (typeof IGNORE_SPACES !== "undefined" && IGNORE_SPACES) {
      if (
        searchTerm
          .split(" ")
          .join("")
          .indexOf(positiveKeyword.split(" ").join("")) > -1
      ) {
        return true;
      }
    }
    return false;
  }

  /**
   * Check if the exact match positive keyword e.g. [dog collar]
   * is in the searchTerm
   * @param {string} positiveKeyword
   * @param {string} searchTerm
   * @returns {boolean}
   */
  _checkExactMatch(positiveKeyword, searchTerm) {
    const wordInString = (s, word) => new RegExp('\\b' + word + '\\b', 'i').test(s);
    positiveKeyword = positiveKeyword.replace("[", "").replace("]", "").trim()
    return wordInString(searchTerm, positiveKeyword)
  }

  _checkRegexMatch(positiveKeyword, searchTerm) {
    //a regex positive keyword migth be regex(.*dog.*collar.*)
    const regex = positiveKeyword.replace("regex(", "").replace(")", "")
    const result = new RegExp(regex, 'i').test(searchTerm)
    // testLog(`checking regex match: ${regex} against ${searchTerm}. Result: ${result}`);
    return result;
  }

  /**
   * we'll allow exact match terms defined like [this]
   * these will match the whole word exactly
   * @param {string} positiveKeyword 
   * @returns whether the word is exact match
   */
  _isExactMatchPositiveWord(positiveKeyword) {
    if (positiveKeyword.indexOf("[") === -1) return false
    if (positiveKeyword.indexOf("]") === -1) return false
    return true
  }


}




/**
 * Take a negative keyword and check if it's necessary
 * i.e. return true if it's already blocked by the existing negatives
 * Args
 * @String negativeKeyword
 * @String negativeKeywordMatchType
 * @Object existingNegatives
 * @return bool
 */
function keywordIsBlocked(negativeKeyword, negativeKeywordMatchType, existingNegatives) {
  //return true if the negative blocks the keyword
  function negativeIsBlockingQuery(query, matchType, negativeKeyword) {
    if (negativeKeyword.match_type == "exact") {
      return addMatchType(query, matchType) == addMatchType(negativeKeyword.keyword, negativeKeyword.match_type);
    }

    //broad match can be in any order but all words must be present
    //singe word phrase behaves the same as broad
    if (
      negativeKeyword.match_type == "broad" ||
      (negativeKeyword.match_type == "phrase" &&
        negativeKeyword.keyword.split(" ").length == 1)
    ) {
      let negativeWords = negativeKeyword.keyword.split(" ");
      let matches = 0;
      for (let negativeWords_i in negativeWords) {
        let negativeWord = negativeWords[negativeWords_i];
        if (wholeWordMatch(negativeWord, query)) matches++;
      }
      if (matches == negativeWords.length) return true;
      return false;
    }

    //phrase match needs to be in order
    //the entire keyword needs to be present
    if (
      negativeKeyword.match_type == "phrase" &&
      negativeKeyword.keyword.split(" ").length > 1
    ) {
      let negativeWords = negativeKeyword.keyword.split(" ");
      let matches = [];
      for (let negativeWords_i in negativeWords) {
        let negativeWord = negativeWords[negativeWords_i];
        if (wholeWordMatch(negativeWord, query))
          matches.push(negativeWord);
      }
      if (
        matches.length == negativeWords.length &&
        query.indexOf(matches.join(" ")) > -1
      )
        return true;
      return false;
    }

    log("negative keyword match type not recognised");
  }

  /*
   * whole match only
   * Args
   * String needle
   * String haystack
   * @return bool
   */
  function wholeWordMatch(needle, haystack) {
    return haystack.split(" ").indexOf(needle) > -1;
  }

  for (let existingNegatives_i in existingNegatives) {
    const existingNegative = existingNegatives[existingNegatives_i];
    if (
      negativeIsBlockingQuery(
        negativeKeyword,
        negativeKeywordMatchType,
        existingNegative
      )
    ) {
      return true;
    }
  }
  return false;
}

function addMatchType(word, matchType) {
  word = String(word);
  if (matchType.toLowerCase() == "broad") {
    word = word.trim();
  } else if (matchType.toLowerCase() == "bmm") {
    word = word
      .split(" ")
      .map(function (x) {
        return "+" + x;
      })
      .join(" ")
      .trim();
  } else if (matchType.toLowerCase() == "phrase") {
    word = '"' + word.trim() + '"';
  } else if (matchType.toLowerCase() == "exact") {
    word = "[" + word.trim() + "]";
  } else {
    throw "Error: Match type not recognised. Please provide one of Broad, BMM, Exact or Phrase";
  }
  return word;
}

/**
 * 
 * @param {string} string 
 * @returns {array} - ['string1', 'string2', 'string3'...]
 */
function stringToCsv(string) {
  if (string.trim() == "") return [];
  return string
    .trim()
    .split(",")
    .map(function (x) {
      return x.trim().replace('\n', '');
    });
}


/**
 * To make life easier
 * We'll add Campaign and Ad Group data to a 
 * 'Ad Group Data' sheet
 * So it can be copied and pasted
 * (If the sheet doesn't exist, nothing will happen)
 */
class AddAdGroupData {

  constructor() {
    this.reportData = {
      name: 'Ad Group Name',
      fields: ['campaign.id', 'campaign.name', 'ad_group.id', 'ad_group.name',],
      filterOnDate: false,
      reportName: 'ad_group',
      filters: [
        { field: 'ad_group.status', operator: '=', value: 'ENABLED' }
      ],
    }
  }

  add() {
    const adGroupDataSheet = SPREADSHEET.getSheetByName('Ad Group Data');
    if (!adGroupDataSheet) {
      return;
    }
    const query = new QueryBuilder(this.reportData.reportName)
      .withFields(this.reportData.fields).withFilters(this.reportData.filters).withGlobalFilters().get();
    const report = AdsApp.report(query);
    report.exportToSheet(adGroupDataSheet)
  }

}


class QueryBuilder {

  constructor(reportName) {
    this.reportName = reportName;
    this.query = 'SELECT ';
  }

  withFields(fields) {
    this.query += fields.join(', ');
    this.query += ` FROM ${this.reportName}`;
    return this;
  }

  withFilters(filters) {
    for (let filter of filters) {
      const operator = this.query.includes('WHERE') ? 'AND' : 'WHERE';
      if (typeof filter.value != 'string') {
        throw new Error(`Filter value ${filter.value} must be a string, got ${typeof filter.value}`);
      }
      this.query += ` ${operator} ${filter.field} ${filter.operator} ${filter.value}`;
    }
    return this;
  }

  withGlobalFilters() {
    if (typeof GLOBAL_FILTERS === 'undefined') {
      return this;
    }
    this.withFilters(GLOBAL_FILTERS)
    return this;
  }

  duringDays(days) {
    const fromDate = this._getDateRange(days);
    const toDate = this._getDateRange(1);
    const operator = this.query.includes('WHERE') ? 'AND' : 'WHERE';
    this.query += ` ${operator} segments.date BETWEEN '${fromDate}' AND '${toDate}'`;
    return this;
  }

  _getDateRange(days) {
    //YYYY-MM-DD
    let date = new Date();
    date = new Date(date.setDate(date.getDate() - days));
    const year = date.getFullYear();
    const month = date.getMonth() < 9 ? `0${date.getMonth() + 1}` : date.getMonth() + 1;
    const day = date.getDate() < 10 ? `0${date.getDate()}` : date.getDate();
    return `${year}-${month}-${day}`;
  }

  get() {
    return this.query;
  }

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


function log(message) {
  if (typeof message === 'object') {
    message = JSON.stringify(message);
  }
  console.log(message);
}


function testLog(message) {
  if (typeof TEST_MODE === 'undefined') {
    return;
  }
  if (!TEST_MODE) {
    return;
  }
  if (typeof message === 'object') {
    message = JSON.stringify(message);
  }
  console.log(message);
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

function fuzzyMatchScore(needle, haystack) {
  let a = FuzzySet();
  a.add(haystack);
  let result = a.get(needle)
  if (!result) return 0
  return result[0][0];
}

