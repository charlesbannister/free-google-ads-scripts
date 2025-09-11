/**
 * Auto Negative Keyword Script (Admin Panel Version)
 * @author Charles Bannister (https://www.linkedin.com/in/charles-bannister/)
 * This script is available as a public GitHub script: https://gist.github.com/charlesbannister/afa9cd70d3b4a010e4fc6ef5d9ae1c3b
 * Google Sheet Template: https://docs.google.com/spreadsheets/d/1xwdogT2HiHV_Gx9pdY4rErhLSm7Qrooj6wbikEjf-G0/edit?gid=0#gid=0
 * Setup your rules at https://autoneg.shabba.io
 */

// --- Configuration ---
// <<<< PASTE SPREADSHEET URL HERE >>>>
const SPREADSHEET_URL = "YOUR_SPREADSHEET_URL_HERE";
const SCRIPT_VERSION = 12;

// <<< /START/ These should all be false >>>
const DEBUG_MODE = false;
const OBFUSCATE_SEARCH_TERMS = false;
const TEST_MODE = false;
// <<< These should all be false /END/>>>

const API_BASE_URL = "https://autoneg.shabba.io/api/rules?sheetId=";
const EXAMPLE_TERMS_API_BASE_URL = "https://autoneg.shabba.io/api/example-terms?sheetId=";
const LATEST_SCRIPT_VERSION_API_URL = "https://autoneg.shabba.io/api/script-info/latest-script-version";
const MESSAGE_API_URL = "https://autoneg.shabba.io/api/script-info/message";
const ADMIN_PANEL_BASE_URL = "https://autoneg.shabba.io/?id=";

const LLM_RESPONSE_DELAY_TIME_MILLISECONDS = 100;

// --- Settings Sheet Configuration ---
const SETTINGS_SHEET_NAME = "Settings";
const SETTINGS_EMAIL_RANGE = "B5";
const SETTINGS_LOG_PROMPTS_RANGE = "B12";
// --- End Configuration ---

let FIRST_PROMPT_LOGGED = false;
// --- Rule Output Sheet Range Configuration ---
const RULE_NAME_RANGE = "A1";
const RULE_LAST_RUNTIME_LABEL_RANGE = "A2";
const RULE_LAST_RUNTIME_VALUE_RANGE = "A3";
const RULE_LAST_EMAIL_LABEL_RANGE = "B2";
const RULE_LAST_EMAIL_VALUE_RANGE = "B3";
const RULE_EDIT_LINK_RANGE = "A4";
const RULE_DATA_START_ROW = 5;
// --- End Rule Output Sheet Configuration ---



const NEGATIVE_MATCH_TYPE = "EXACT";

const CAMPAIGN_TYPES = {
  "shopping": "Shopping",
  "search": "Search",
  "pmax": "Performance Max"
}

const DEFAULT_CLAUDE_MODEL_NAME = "claude-3-5-sonnet-20240620";
const DEFAULT_AWAN_MODEL_NAME = 'Meta-Llama-3.1-8B-Instruct';
const DEFAULT_CHAT_GPT_MODEL_NAME = 'gpt-4o-mini';
const DEFAULT_GEMINI_MODEL_NAME = 'gemini-1.5-flash';
const DEFAULT_MODEL_NAMES = {
  CLAUDE: DEFAULT_CLAUDE_MODEL_NAME,
  AWAN: DEFAULT_AWAN_MODEL_NAME,
  CHAT_GPT: DEFAULT_CHAT_GPT_MODEL_NAME,
  GEMINI: DEFAULT_GEMINI_MODEL_NAME,
}

const LLM_NAME_ENUM = Object.freeze({
  CLAUDE: 'CLAUDE',
  AWAN: 'AWAN',
  CHAT_GPT: 'CHAT_GPT',
  GEMINI: 'GEMINI',
});



//API metric to header mapping
const BASE_METRICS = {
  "metrics.impressions": "Impressions",
  "metrics.clicks": "Clicks",
  "metrics.cost": "Cost",//"cost" is the correct API metric, but "cost_micros" is the correct field
  "metrics.conversions": "Conversions",
  "metrics.conversions_value": "Conv. Value",//note the API calls this "conversion_value" (no s in conversions)
}

const API_PERCENTAGE_METRICS = [
  "conversion_rate",
  "ctr",
  "roas",
]


function main() {
  console.log('Started');
  VersionLogger.logScriptVersionMessage();
  MessageLogger.logMessage();
  if (OBFUSCATE_SEARCH_TERMS) {
    console.warn("WARNING: OBFUSCATE_SEARCH_TERMS is true. This will obfuscate search terms in the output.");
  }
  if (!isMCC()) {
    runAccount();
    return;
  }
  if (typeof ACCOUNT_ID === 'undefined') {
    console.error("To run at MCC/Manager level, specify an ACCOUNT_ID in the script.");
    return;
  }
  MccApp.accounts()
    .withIds([ACCOUNT_ID])
    .withLimit(50)
    .executeInParallel("runAccount");
}

const VersionLogger = {

  logScriptVersionMessage() {
    try {
      const latestVersion = this.getLatestScriptVersion();
      if (latestVersion > SCRIPT_VERSION) {
        const sheetId = extractSpreadsheetId_(SPREADSHEET_URL);
        console.warn(`A new version of the script is available. You're on version ${SCRIPT_VERSION}. \nGet version ${latestVersion} at ${ADMIN_PANEL_BASE_URL}${sheetId}`);
      }
    } catch (e) {
      // throw new Error(`ERROR fetching latest script version: ${e.message}`);
    }
  },

  getLatestScriptVersion() {
    const url = LATEST_SCRIPT_VERSION_API_URL;
    const response = UrlFetchApp.fetch(url);
    const responseBody = response.getContentText();
    const data = JSON.parse(responseBody);
    return data["version"];
  },

}

const MessageLogger = {
  logMessage() {
    try {
      const url = MESSAGE_API_URL;
      const response = UrlFetchApp.fetch(url);
      const responseBody = response.getContentText();
      const data = JSON.parse(responseBody);
      if (!data["message"]) {
        return;
      }
      console.log(data["message"]);
    } catch (e) {
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


/**
 * Main function executed by Google Ads Scripts.
 */
function runAccount() {
  if (SPREADSHEET_URL === "YOUR_SPREADSHEET_URL_HERE") {
    console.log("ERROR: Please replace 'YOUR_SPREADSHEET_URL_HERE' with your actual Google Sheet URL in the script.");
    return;
  }

  const spreadsheetId = extractSpreadsheetId_(SPREADSHEET_URL);
  if (!spreadsheetId) {
    console.log("ERROR: Could not extract Spreadsheet ID from the provided URL: " + SPREADSHEET_URL);
    return;
  }

  console.log(`Spreadsheet Url: ${SPREADSHEET_URL}`);
  console.log(`Admin Panel Url: ${ADMIN_PANEL_BASE_URL + spreadsheetId}\n`);
  const apiUrl = API_BASE_URL + spreadsheetId;
  console.log("Fetching rules from: " + apiUrl);

  let rules = fetchRules_(apiUrl);
  rules = rules.filter(rule => rule.enabled);

  for (let rule of rules) {
    console.log(`${rule.name} - ${rule.id} - enabled: ${rule.enabled}`);
  }

  if (!rules || rules.length === 0) {
    console.log("No rules found or error fetching rules. Exiting.");
    return;
  }

  console.log("Fetched " + rules.length + " rules. Processing...");

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);

  // Get global email recipients from Settings sheet
  const globalEmailRecipients = getGlobalEmailRecipients_(spreadsheet);
  if (globalEmailRecipients.length > 0) {
    console.log(`Found ${globalEmailRecipients.length} global email recipient(s) in Settings sheet: ${globalEmailRecipients.join(", ")}`);
  }

  // Sort rules by last runtime
  rules = sortRulesByLastRuntime_(rules, spreadsheet);
  console.log(`Rules sorted by last runtime. Rules with no previous runs will be processed first.`);

  rules.forEach(function (rule) {
    try {
      console.log("\nProcessing rule: '" + rule.name);
      processRule_(rule, spreadsheet, globalEmailRecipients);
      console.log("Finished processing rule: '" + rule.name + "'");
    } catch (e) {
      console.log("ERROR processing rule '" + rule.name + "' (ID: " + rule.id + "): " + e.stack);
      console.error(e.stack)
      throw e;
      // Optional: Log error to a specific sheet or send notification
    }
  });

  console.log("Script finished.");
}

/**
 * Fetches rules from the specified API URL, managing retries with exponential backoff for 502 errors.
 * @param {string} apiUrl The URL to fetch rules from.
 * @return {Array<Object>|null} An array of rule objects or null if an error occurs after retries.
 * @private
 */
function fetchRules_(apiUrl) {
  const MAX_RETRIES = 2; // Number of retries after the initial attempt
  const INITIAL_DELAY_MS = 5 * 1000; // 5 seconds
  const BACKOFF_FACTOR = 4; // Delay increases by this factor (5s, 20s)
  let currentDelayMs = INITIAL_DELAY_MS;
  let attempts = 0;

  while (attempts <= MAX_RETRIES) {
    attempts++;
    console.log(`Fetch rules attempt ${attempts} from: ${apiUrl}`);

    const result = _fetchAndParseRulesAttempt_(apiUrl);

    if (result.success) {
      console.log(`Successfully fetched ${result.rules.length} rules on attempt ${attempts}.`);
      return result.rules; // Success!
    }

    // --- Handle Failures ---
    const canRetry = (result.errorType === 'http' && result.responseCode === 502 && attempts <= MAX_RETRIES);

    if (canRetry) {
      console.log(`Received HTTP 502 (Attempt ${attempts}). Retrying in ${currentDelayMs / 1000} seconds...`);
      Utilities.sleep(currentDelayMs);
      currentDelayMs *= BACKOFF_FACTOR; // Increase delay for next potential retry
    } else {
      // Log final failure details
      console.log(`Failed to fetch rules on attempt ${attempts}. Error Type: ${result.errorType}.`);
      if (result.responseCode) console.log(`  HTTP Status: ${result.responseCode}`);
      if (result.message) console.log(`  Message: ${result.message}`);
      if (result.responseBody) console.log(`  Response Body: ${result.responseBody}`); // Log response body on final HTTP failure
      return null; // Indicate final failure
    }
  } // End while loop

  // This point should only be reached if all retries failed (specifically on 502s)
  console.log("Exhausted all fetch attempts after encountering 502 errors.");
  return null;
}

/**
 * Performs a single attempt to fetch rules from the API, parse the JSON, and validate the structure.
 * @param {string} apiUrl The URL to fetch rules from.
 * @return {Object} An object indicating the outcome:
 *   - { success: true, rules: Array<Object> } on success.
 *   - { success: false, errorType: string, responseCode?: number, responseBody?: string, message?: string } on failure.
 *     Error types: 'http', 'parse', 'validation', 'fetch'.
 * @private
 */
function _fetchAndParseRulesAttempt_(apiUrl) {
  try {
    const options = {
      method: 'get',
      contentType: 'application/json',
      muteHttpExceptions: true // Allows checking response code
    };
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      // Success Case
      let rules;
      try {
        rules = JSON.parse(responseBody);
      } catch (parseError) {
        console.log(`ERROR parsing JSON response. Body: ${responseBody}. Error: ${parseError}`);
        return { success: false, errorType: 'parse', message: parseError.message, responseCode: responseCode };
      }

      if (!Array.isArray(rules)) {
        console.log(`ERROR: API response was not an array. Body: ${responseBody}`);
        return { success: false, errorType: 'validation', message: 'Response was not an array.', responseCode: responseCode };
      }
      // console.log("Single fetch attempt successful.");
      return { success: true, rules: rules };

    } else {
      // HTTP Error Case
      // console.log(`HTTP Error during fetch attempt: ${responseCode}`);
      return { success: false, errorType: 'http', responseCode: responseCode, responseBody: responseBody, responseCode };
    }

  } catch (fetchError) {
    // Network/Fetch Exception Case
    console.log(`Exception during UrlFetchApp.fetch: ${fetchError}`);
    return { success: false, errorType: 'fetch', message: fetchError.message };
  }
}


/**
 * Processes a single rule: parses conditions, builds query, fetches report,
 * checks conditions, optionally applies negative keywords, and writes report to sheet.
 * @param {Object} rule The rule object.
 * @param {Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @param {Array<string>} globalEmailRecipients Optional global email recipients from Settings sheet.
 * @private
 */
function processRule_(rule, spreadsheet, globalEmailRecipients = []) {
  const ruleLevel = (rule.level || 'campaign').toLowerCase();
  const autoApply = rule.auto_apply === true;
  const generateReport = rule.generate_report === true;
  const sendEmail = rule.send_email === true;
  const maxSearchTerms = rule.max_search_terms ? parseInt(rule.max_search_terms) : null;
  const searchTermStatusParams = {
    searchTermStatusAddedExcluded: rule.search_term_status_added_excluded,
    searchTermStatusAdded: rule.search_term_status_added,
    searchTermStatusExcluded: rule.search_term_status_excluded
  }

  // --- 0. Check if any action is needed ---
  if (!generateReport && !autoApply) {
    console.log(`Neither report generation nor auto-apply is enabled for rule "${rule.name}". Skipping.`);
    return;
  }

  // --- 1. Parse Conditions ---
  const conditions = _parseRuleConditions_(rule);
  if (!conditions) return; // Stop if parsing failed

  // --- 2. Build GAQL Query ---
  let gaqlQuery;
  try {
    gaqlQuery = buildGaqlQuery_(ruleLevel, rule.lookback_days, conditions.entityConditions, maxSearchTerms, conditions.performanceConditions, searchTermStatusParams);
  } catch (e) {
    console.log(`ERROR building GAQL query for rule "${rule.name}": ${e.message}. Skipping rule.`);
    return;
  }

  // --- 3. Execute Report ---
  const regularRows = _executeReport_(gaqlQuery, rule.name, maxSearchTerms);

  if (DEBUG_MODE) {
    writeResultsToDebugSheet(regularRows);
  }

  if (!regularRows) return; // Stop if report failed

  // --- 3b. Try to get PMAX data if this is a campaign-level rule ---
  let pmaxRows = [];
  if (ruleLevel === 'campaign') {
    pmaxRows = _fetchPmaxData_(rule.name, rule.lookback_days, conditions.entityConditions);
  } else {
    console.log(`Rule "${rule.name}" is at ${ruleLevel} level. PMAX search terms will only be fetched for campaign-level rules.`);
  }

  // --- 3c. Get example terms ---
  const exampleTerms = rule.include_examples ? _fetchExampleTerms_(spreadsheet.getId(), rule.id, regularRows, ruleLevel) : [];
  // console.log(`Example terms: ${JSON.stringify(exampleTerms, null, 2)}`);


  // --- 4. Process Rows (Check conditions, collect data/negatives) ---
  const processingResult = _processReportRows_({
    regularRows: regularRows,
    ruleLevel: ruleLevel,
    autoApply: autoApply,
    generateReport: generateReport,
    performanceConditions: conditions.performanceConditions,
    textConditions: conditions.textConditions,
    pmaxRows: pmaxRows,
    rule: rule,
    exampleTerms: exampleTerms
  });

  console.log(`Finished processing rows for rule "${rule.name}". Checked: ${processingResult.stats.checked}, Matched Filters: ${processingResult.stats.matched}, Negatives Collected: ${processingResult.stats.collectedForNeg}.`);

  // --- 5. Auto-Apply Negatives (if enabled and not in TEST_MODE) ---
  if (autoApply) {
    if (TEST_MODE && !AdsApp.getExecutionInfo().isPreview()) {
      console.log(`--> TEST_MODE is enabled. Skipping actual negative keyword application for rule "${rule.name}".`);
      // Optional: Log counts of what would be added per entity
      if (ruleLevel === 'adgroup') {
        console.log(`--> TEST MODE: Ad Group Negatives Collected: ${JSON.stringify(processingResult.negativesByAdGroup, null, 2)}`);
      } else if (ruleLevel === 'campaign') {
        console.log(`--> TEST MODE: Campaign Negatives Collected: ${JSON.stringify(processingResult.negativesByCampaign, null, 2)}`);
      }
    } else {
      // Apply the collected negatives
      _applyNegativeKeywords_(
        rule.name,
        ruleLevel,
        processingResult.negativesByAdGroup,
        processingResult.negativesByCampaign,
        processingResult.negativesByAccount,
        rule.negative_list_name
      );
    }
  }

  // --- 6. Write Report to Sheet (if enabled) ---
  let sheetUrl = '';
  const ruleSheet = _getOrCreateSheet_(spreadsheet, rule.rule_number);
  const emailLastSentTimestamp = ruleSheet.getRange(RULE_LAST_EMAIL_VALUE_RANGE).getValue();
  if (generateReport) {
    let headers = [
      "Search Term", "Added/Removed",
      "Impressions", "Clicks", "Cost", "Conversions", "Conv. Value",
      "CTR", "Conv. Rate", "Cost/Conv.", "CPC", "ROAS", "Negative Added", "AI Response"
    ];
    if (ruleLevel === 'campaign' || ruleLevel === 'adgroup') {
      headers.splice(1, 0, "Campaign Name");
    }
    if (ruleLevel === 'adgroup') {
      headers.splice(2, 0, "Ad Group Name");
    }
    // console.log(`headers: ${JSON.stringify(headers, null, 2)}`);
    // console.log(`processingResult.outputDataRows: ${JSON.stringify(processingResult.outputDataRows, null, 2)}`);
    const outputData = [headers].concat(processingResult.outputDataRows); // Combine headers and data rows
    writeToSheet_(spreadsheet, rule.name, outputData, rule.rule_number, rule.id, ruleSheet);

    // Get the sheet URL for email notifications
    const spreadsheetId = extractSpreadsheetId_(SPREADSHEET_URL);
    const formattedSheetName = isNumeric(rule.rule_number) ? `${rule.rule_number}` : String(rule.rule_number);
    sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${ruleSheet.getSheetId()}`;
  } else {
    console.log(`Report generation skipped for rule "${rule.name}" as generate_report was not true.`);
  }

  // --- 7. Send Email Notification (if enabled) ---
  rule.numberOfResults = processingResult.outputDataRows.length;
  if (sendEmail && rule.numberOfResults > 0) {
    const isPreview = AdsApp.getExecutionInfo().isPreview();
    const isTestMode = TEST_MODE && !isPreview;

    // Generate execution mode string
    let executionMode = "Live Mode";
    if (isPreview) executionMode = "Preview Mode";
    if (isTestMode) executionMode = "Test Mode";


    // Create and send email
    const emailService = new EmailService({
      rule: rule,
      processingResult: processingResult,
      reportUrl: sheetUrl,
      executionMode: executionMode,
      ruleLevel: ruleLevel,
      timestamp: Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      accountName: AdsApp.currentAccount().getName(),
      accountId: AdsApp.currentAccount().getCustomerId(),
      globalEmailRecipients: globalEmailRecipients,
      emailLastSentTimestamp,
      ruleSheet
    });
    emailService.sendEmail();
  }
}

/**
 * 
  * @param {string} spreadsheetId The ID of the Google Spreadsheet
  * @param {string} ruleId The ID of the rule being processed
  * @param {Object} regularRows Iterator containing search term data rows
  * @param {string} ruleLevel The level at which the rule operates ('account' or 'adgroup')
  * @returns {Array<Object>} Array of example search terms with campaign/ad group data
 */
function _fetchExampleTerms_(spreadsheetId, ruleId, regularRows, ruleLevel) {
  const url = `${EXAMPLE_TERMS_API_BASE_URL}${spreadsheetId}&ruleId=${ruleId}`;
  console.log(`Fetching example terms from: ${url}`);
  let response;
  try {
    response = UrlFetchApp.fetch(url);
  } catch (error) {
    console.log(`ERROR fetching example terms: ${error.message}`);
    return []; // Return an empty array on error
  }
  const responseBody = response.getContentText();
  const data = JSON.parse(responseBody);
  const exampleTerms = data['exampleTerms'];
  const exampleTermsArray = exampleTerms.map(term => term.term);

  // Create a copy of regularRows without consuming the original iterator
  const regularRowsItems = [];
  while (regularRows.hasNext()) {
    regularRowsItems.push(regularRows.next());
  }

  // Create a new iterator for processing example terms
  const regularRowsCopy = createIteratorFromItems(regularRowsItems);

  // Recreate the original iterator for later use
  const originalIteratorRestored = createIteratorFromItems(regularRowsItems);

  // Replace the consumed regularRows with our restored version
  Object.assign(regularRows, originalIteratorRestored);

  // console.log(`Regular rows: ${JSON.stringify(regularRowsItems, null, 2)}`);

  if (ruleLevel === 'account' || !regularRowsCopy.hasNext()) {
    return exampleTermsArray.map(term => ({
      term: term,
      campaignName: null,
      campaignId: null,
      adGroupName: null,
      adGroupId: null
    }));
  }
  let entityData = [];
  // get campaign name, id and ad group name, id from regularRows
  while (regularRowsCopy.hasNext()) {
    const row = regularRowsCopy.next();
    entityData.push({
      campaignName: row['campaign.name'],
      campaignId: row['campaign.id'],
      adGroupName: row['ad_group.name'],
      adGroupId: row['ad_group.id']
    });
  }
  entityData = [...new Set(entityData.map(JSON.stringify))].map(JSON.parse);

  // for each example term, create a row for each campaign or ad group
  let exampleTermRows = [];
  for (exampleTerm of exampleTermsArray) {
    for (entity of entityData) {
      exampleTermData = {
        term: exampleTerm,
        campaignName: entity.campaignName,
        campaignId: entity.campaignId,
        adGroupName: entity.adGroupName,
        adGroupId: entity.adGroupId,
        termType: 'EXAMPLE'
      }
      // add only if the object isn't already in the array
      const found = exampleTermRows.find(row => row.term === exampleTermData.term && row.campaignName === exampleTermData.campaignName && row.adGroupName === exampleTermData.adGroupName);
      if (!found) {
        exampleTermRows.push(exampleTermData);
      }
    }
  }

  return exampleTermRows;
}

/**
 * Creates a new iterator from an array of items
 * @param {Array} items Array of items to iterate over
 * @return {Object} An iterator object
 */
function createIteratorFromItems(items) {
  let index = 0;
  return {
    hasNext: function () {
      return index < items.length;
    },
    next: function () {
      if (this.hasNext()) {
        return items[index++];
      }
      throw new Error("No more items in the iterator.");
    },
    // Add a reset method to allow reusing the iterator
    reset: function () {
      index = 0;
      return this;
    }
  };
}

function writeResultsToDebugSheet(rows) {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const debugSheet = spreadsheet.getSheetByName('Debug Sheet');
  if (!debugSheet) {
    console.log("Debug sheet not found. Skipping write to debug sheet.");
    return;
  }
  debugSheet.clear();
  debugSheet.getRange("A1").setValue(JSON.stringify(rows));
}

/**
 * Parses the condition JSON strings from a rule object.
 * @param {Object} rule The rule object containing condition strings.
 * @return {{entityConditions: Array, performanceConditions: Array, textConditions: Object}|null}
 *         An object with parsed conditions, or null if parsing fails.
 * @private
 */
function _parseRuleConditions_(rule) {

  try {
    const entityConditions = rule.entity_conditions;
    let performanceConditions = rule.performance_conditions;
    performanceConditions = performanceConditions.filter(value => value !== "");
    const textConditions = rule.text_conditions;
    // Basic validation (can be expanded)
    if (!Array.isArray(entityConditions) || !Array.isArray(performanceConditions) || typeof textConditions !== 'object') {
      throw new Error("Invalid condition format after parsing.");
    }
    // in the performance conditions, replace conversion_value with conversions_value
    performanceConditions = performanceConditions.map(condition => {
      // we named conversions_value "conversion_value" in the web app (API))
      // so we'll rename it here
      if (condition.metric === 'conversion_value') {
        condition.metric = 'conversions_value';
      }
      return condition;
    });

    //divide API_PERCENTAGE_METRICS by 100
    performanceConditions = performanceConditions.map(condition => {
      if (API_PERCENTAGE_METRICS.includes(condition.metric)) {
        condition.value = parseFloat(condition.value) / 100;
      }
      return condition;
    });

    console.log(`Successfully parsed conditions for rule "${rule.name}"`);
    return { entityConditions, performanceConditions, textConditions };
  } catch (e) {
    console.log(`ERROR parsing conditions for rule "${rule.name}": ${e}.`);
    console.error(e.stack)
    return null;
  }
}

/**
 * Executes the GAQL query and returns the report rows iterator.
 * @param {string} gaqlQuery The GAQL query string.
 * @param {string} ruleName For logging context.
 * @param {number|null} maxSearchTerms Optional maximum number of search terms in the rule.
 * @return {Iterator|null} The report rows iterator or null if execution fails.
 * @private
 */
function _executeReport_(gaqlQuery, ruleName, maxSearchTerms = null) {
  try {
    console.log(`Executing report for rule "${ruleName}"...`);
    const report = AdsApp.report(gaqlQuery);

    const rows = report.rows();
    const totalEntities = rows.totalNumEntities();
    // Warn if there might be more data than what's being returned
    if (maxSearchTerms && totalEntities >= maxSearchTerms) {
      console.log(`WARNING: Rule "${ruleName}" found ${totalEntities} search terms but is limited to ${maxSearchTerms}. ` +
        `Some search terms will not be processed.`);
    }

    return rows;
  } catch (e) {
    console.log(`ERROR executing report for rule "${ruleName}" with query [${gaqlQuery}]: ${e}.`);
    return null;
  }
}

/**
 * Processes report rows: checks conditions, collects data for reporting and negatives.
 * @param {Object} params Object containing:
 * @param {Iterator} params.regularRows The report rows iterator.
 * @param {string} params.ruleLevel 'campaign' or 'adgroup'.
 * @param {boolean} params.autoApply Whether to collect negatives for auto-application.
 * @param {boolean} params.generateReport Whether to collect data for the report sheet.
 * @param {Array} params.performanceConditions Parsed performance conditions.
 * @param {Object} params.textConditions Parsed text conditions.
 * @param {Array} params.pmaxRows Optional array of PMAX search term rows.
 * @param {Object} params.rule The rule object.
 * @param {Array} params.exampleTerms Array of example terms.
 * @return {{outputDataRows: Array<Array>, negativesByAdGroup: Object, negativesByCampaign: Object, stats: {checked: number, matched: number, collectedForNeg: number}}}
 *         Collected data and stats.
 */
function _processReportRows_(params) {
  const { regularRows, ruleLevel, autoApply, generateReport, performanceConditions, textConditions, pmaxRows, rule, exampleTerms } = params;

  // Initialize result containers
  const result = {
    outputDataRows: [],
    stats: { checked: 0, matched: 0, collectedForNeg: 0 }
  };

  // Get execution mode information
  const isPreview = AdsApp.getExecutionInfo().isPreview();
  const isTestMode = TEST_MODE && !isPreview;

  // Status messages for the "Negative Added" column
  const STATUS = {
    ADDED: "Yes",
    PREVIEWED: "Yes (Previewed)",
    AUTO_APPLY_DISABLED: "No (Auto Apply Disabled)",
    TEST_MODE: "No (TEST_MODE enabled)",
    PMAX_NOT_SUPPORTED: "PMAX not supported"
  };

  // Process PMAX search term rows
  _processPmaxRows_(pmaxRows, ruleLevel, autoApply, generateReport, textConditions, STATUS, result, rule);

  result.negativesByAdGroup = {};
  result.negativesByCampaign = {};
  result.negativesByAccount = {};

  // Process regular search term rows
  _processRegularRows_(regularRows, ruleLevel, autoApply, generateReport, performanceConditions, textConditions, STATUS, result, rule);

  // Process example terms
  console.log(`\nProcessing example terms`);
  _processExampleTerms_(exampleTerms, ruleLevel, autoApply, generateReport, textConditions, STATUS, result, rule);

  return result;
}

/**
 * Process regular search term rows from the report.
 * @param {Iterator} rows The report rows iterator.
 * @param {string} ruleLevel 'campaign' or 'adgroup'.
 * @param {boolean} autoApply Whether to collect negatives for auto-application.
 * @param {boolean} generateReport Whether to collect data for the report sheet.
 * @param {Array} performanceConditions Parsed performance conditions.
 * @param {Object} textConditions Parsed text conditions.
 * @param {Object} STATUS Status message constants.
 * @param {Object} result The result object to update.
 * @param {Object} rule The rule object.
 * @private
 */
function _processRegularRows_(rows, ruleLevel, autoApply, generateReport, performanceConditions, textConditions, STATUS, result, rule) {
  console.log(`Processing report rows for rule level: ${ruleLevel}`);
  if (!rows || !rows.hasNext()) {
    console.log(`No rows found for rule "${rule.name}". Skipping.`);
    return;
  }

  console.log(`Processing ${rows.totalNumEntities()} regular search term rows`);

  while (rows.hasNext()) {
    const row = rows.next();
    // Add a query_source property to identify this row comes from the regular GAQL query
    row.query_source = 'GAQL';
    result.stats.checked++;

    // Extract basic info
    const searchTerm = OBFUSCATE_SEARCH_TERMS ? obfuscateSearchTerm(row['search_term_view.search_term']) : row['search_term_view.search_term'];
    const campaignId = row['campaign.id'];
    const campaignName = row['campaign.name'];
    const addedRemoved = row['search_term_view.status'];

    if (DEBUG_MODE) {
      console.log(`Processing search term row. search term: ${searchTerm}, campaign name: ${campaignName},   addedRemoved: ${addedRemoved}`);
    }

    let adGroupId = null;
    let adGroupName = null;

    if (ruleLevel === 'adgroup') {
      if (typeof row['ad_group.id'] === 'undefined' || typeof row['ad_group.name'] === 'undefined') {
        console.log(`Warning: ad_group.id/name missing in adgroup-level rule row. Term: "${searchTerm}". Skipping.`);
        continue;
      }
      adGroupId = row['ad_group.id'];
      adGroupName = row['ad_group.name'];
    }

    // Skip if row doesn't meet conditions
    if (!_rowMeetsAllConditions_(row, searchTerm, performanceConditions, textConditions, rule, campaignName, adGroupName, campaignId, adGroupId)) {
      continue;
    }

    // Process matching terms
    result.stats.matched++;
    processMatchedTerm_({
      searchTerm: searchTerm,
      campaignId: campaignId,
      campaignName: campaignName,
      adGroupId: adGroupId,
      adGroupName: adGroupName,
      addedRemoved: addedRemoved,
      row: row,
      ruleLevel: ruleLevel,
      autoApply: autoApply,
      generateReport: generateReport,
      STATUS: STATUS,
      negativesToAddByAdGroup: result.negativesByAdGroup,
      negativesToAddByCampaign: result.negativesByCampaign,
      negativesToAddByAccount: result.negativesByAccount,
      outputDataRows: result.outputDataRows,
      termType: 'STANDARD',
      aiResponseString: row.aiResponseString || "N/A"
    });

    if (autoApply) result.stats.collectedForNeg++;
  }
}

/**
 * Process PMAX search term rows.
 * @param {Array} pmaxRows Array of PMAX search term rows.
 * @param {string} ruleLevel 'campaign' or 'adgroup'.
 * @param {boolean} autoApply Whether to collect negatives for auto-application.
 * @param {boolean} generateReport Whether to collect data for the report sheet.
 * @param {Object} textConditions Parsed text conditions.
 * @param {Object} STATUS Status message constants.
 * @param {Object} result The result object to update.
 * @param {Object} rule The rule object.
 * @private
 */
function _processPmaxRows_(pmaxRows, ruleLevel, autoApply, generateReport, textConditions, STATUS, result, rule) {
  if (ruleLevel !== 'campaign') {
    return;
  }
  if (!pmaxRows || pmaxRows.length === 0) {
    console.log(`No PMAX search term rows found for rule "${rule.name}". Skipping.`);
    return;
  }

  console.log(`Processing ${pmaxRows.length} PMAX search term rows`);

  for (const pmaxRow of pmaxRows) {
    try {
      // Add a query_source property to identify this row is from PMAX
      pmaxRow.query_source = 'PMAX';
      result.stats.checked++;

      // Extract basic info from PMAX row
      const searchTerm = OBFUSCATE_SEARCH_TERMS ? obfuscateSearchTerm(pmaxRow.searchTerm) : pmaxRow.searchTerm;
      const campaignId = pmaxRow.campaignId;
      const campaignName = pmaxRow.campaignName;
      const addedRemoved = "--";

      if (!searchTerm) {
        console.log(`Warning: PMAX row missing search term. Skipping.`);
        continue;
      }

      // Check conditions for PMAX terms
      if (!_pmaxRowMeetsConditions_(searchTerm, textConditions, rule, campaignName, null, campaignId, null)) {
        continue;
      }

      // Process matching terms - PMAX doesn't support adgroups
      result.stats.matched++;
      processMatchedTerm_({
        searchTerm: searchTerm,
        campaignId: campaignId,
        campaignName: campaignName,
        adGroupId: null,
        adGroupName: null,
        addedRemoved: addedRemoved,
        row: pmaxRow,
        ruleLevel: 'campaign',
        autoApply: autoApply,
        generateReport: generateReport,
        STATUS: STATUS,
        negativesToAddByAdGroup: result.negativesByAdGroup,
        negativesToAddByCampaign: result.negativesByCampaign,
        negativesToAddByAccount: result.negativesByAccount,
        outputDataRows: result.outputDataRows,
        termType: 'PMAX',
        aiResponseString: pmaxRow.aiResponseString || "N/A"
      });
      // Don't increment negativesCollectedCount for PMAX terms as they can't be negated
    } catch (pmaxTermError) {
      console.log(`Error processing individual PMAX search term: ${pmaxTermError.message}`);
      // Continue with the next PMAX term
    }
  }
}


function obfuscateSearchTerm(searchTerm) {
  // replace each word in the search term with a random word
  const randomWords = [
    "apple", "banana", "cherry", "date", "elderberry", "fig", "grape", "honeydew",
    "kiwi", "lemon", "mango", "nectarine", "orange", "papaya", "quince", "raspberry",
    "strawberry", "tangerine", "ugli", "vanilla", "watermelon", "xigua", "yam",
    "zucchini", "apricot", "blackberry", "cantaloupe", "dragonfruit", "eggplant",
    "fennel", "garlic", "horseradish", "jalapeno", "kale", "lime", "mushroom",
    "nutmeg", "olive", "peach", "quinoa", "radish", "spinach", "tomato", "ugli",
    "vetch", "wasabi", "xanadu", "yarrow", "zest", "almond", "basil", "cinnamon",
    "dill", "endive", "fennel", "ginger", "herb", "iceberg", "jalapeno", "kohlrabi",
    "leek", "miso", "nori", "oyster", "parsley", "quail", "rhubarb", "sorrel",
    "thyme", "upland", "vermouth", "wasabi", "xanadu", "yarrow", "zucchini",
    "acorn", "beet", "cabbage", "daikon", "eggplant", "fava", "garbanzo", "hops",
    "indigo", "jicama", "kale", "lima", "mangetout", "napa", "okra", "parsnip",
    "quince", "radicchio", "sunchoke", "taro", "upland", "vetch", "watercress",
    "yam", "zucchini",
    "Zeus", "Hera", "Poseidon", "Athena", "Apollo", "Artemis", "Ares", "Aphrodite",
    "Hermes", "Demeter", "Hades", "Persephone", "Dionysus", "Hephaestus", "Hestia",
    "Eros", "Hecate", "Nike", "Pan", "Prometheus", "Atlas", "Medusa", "Cerberus",
    "Chiron", "Theseus", "Hercules", "Odysseus", "Achilles", "Agamemnon", "Helen",
    "Paris", "Cassandra", "Penelope", "Circe", "Calypso", "Narcissus", "Sisyphus",
    "Tantalus", "Orpheus", "Eurydice", "Icarus", "Daedalus", "Midas", "Pygmalion",
    "Ariadne", "Minotaur", "Charybdis", "Scylla", "Furies", "Graces", "Muses",
    "Nymphs", "Satyrs", "Gorgons", "Titans", "Cyclopes", "Hecatoncheires", "Echidna",
    "Typhon", "Rhea", "Cronus", "Gaia", "Uranus", "Eros", "Thanatos", "Hypnos",
    "Nemesis", "Eris", "Hades", "Thanatos", "Asclepius", "Hygieia", "Panacea",
    "Iris", "Hebe", "Eos", "Selene", "Helios", "Nyx", "Chronos", "Ananke",
    "Phobos", "Deimos", "Thanatos", "Elysium", "Styx", "Lethe", "Acheron",
    "Cocytus", "Tartarus", "Chiron", "Aesculapius", "Hades", "Persephone",
    "Elysium", "Fates", "Moirae", "Keres", "Nereids", "Oceanids", "Hesperides",
    "Argentina", "Brazil", "Canada", "Denmark", "Egypt", "France", "Germany",
    "Hungary", "India", "Japan", "Kenya", "Lithuania", "Mexico", "Norway",
    "Oman", "Peru", "Qatar", "Russia", "Spain", "Turkey", "Uganda", "Vietnam",
    "Zimbabwe", "Andes", "Himalayas", "Rockies", "Alps", "Appalachians",
    "Atlas", "Urals", "Pyrenees", "Cascades", "Sierra", "Carpathians",
    "Kilimanjaro", "Aconcagua", "Denali", "Mont Blanc", "Matterhorn",
    "Everest", "Vesuvius", "K2", "Makalu", "Lhotse", "Torre del Paine",
    "Mediterranean Sea", "Caribbean Sea", "Red Sea", "Baltic Sea",
    "North Sea", "Black Sea", "Caspian Sea", "Arabian Sea", "Bering Sea",
    "Coral Sea", "Java Sea", "Philippine Sea", "Tasman Sea", "Weddell Sea",
    "Bay of Bengal", "Gulf of Mexico", "Strait of Gibraltar", "English Channel",
    "Sea of Japan", "Aegean Sea", "Adriatic Sea", "Timor Sea", "Sargasso Sea",
    "South China Sea", "Beaufort Sea", "Hudson Bay", "Labrador Sea",
    "Baffin Bay", "Davis Strait", "Ross Sea", "Amundsen Sea", "Weddell Sea",
    "Seychelles", "Maldives", "Fiji", "Hawaii", "Galapagos", "Bermuda",
    "Iceland", "Greenland", "New Zealand", "Svalbard", "Cuba", "Bahamas"
  ].map(word => word.toLowerCase());

  const doNotObfuscate = ["in",
    "for",
    "on",
    "with",
    "at",
    "by",
    "small",
    "medium",
    "large",
    "xlarge",
    "xxlarge",
    "xxxlarge",
    "xxxxlarge",
    "ikea",
    "of",
    "up", "down", "out", "over", "under", "again", "further", "then", "once", "here", "there", "when", "where", "why", "how", "all", "any", "some", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"];

  // replace each word in the search term with a random word
  const obfuscatedTerm = searchTerm.replace(/\b\w+\b/g, (originalWord) => {
    const word = randomWords[Math.floor(Math.random() * randomWords.length)];
    return doNotObfuscate.includes(originalWord) ? originalWord : word;
  });
  return obfuscatedTerm;
}

/**
 * Process example terms.
 * @param {Array} exampleTerms Array of example terms.
 * @param {string} ruleLevel 'campaign' or 'adgroup'.
 * @param {boolean} autoApply Whether to collect negatives for auto-application.
 * @param {boolean} generateReport Whether to collect data for the report sheet.
 * @param {Object} textConditions Parsed text conditions.
 * @param {Object} STATUS Status message constants.
 * @param {Object} result The result object to update.
 * @param {Object} rule The rule object.
 * @private
 */
function _processExampleTerms_(exampleTerms, ruleLevel, autoApply, generateReport, textConditions, STATUS, result, rule) {
  console.log(`Processing report rows for rule level: ${ruleLevel}`);
  if (!exampleTerms || exampleTerms.length === 0) {
    console.log(`No example terms found for rule "${rule.name}". Skipping.`);
    return;
  }

  console.log(`Processing ${exampleTerms.length} example terms`);

  for (const exampleTerm of exampleTerms) {
    // Add a query_source property to identify this row comes from the regular GAQL query
    result.stats.checked++;

    // Skip if row doesn't meet conditions
    if (!_rowMeetsAllConditions_(exampleTerm, exampleTerm.term, [], textConditions, rule, exampleTerm.campaignName, exampleTerm.adGroupName, exampleTerm.campaignId, exampleTerm.adGroupId)) {
      continue;
    }

    // Process matching terms
    result.stats.matched++;
    processMatchedTerm_({
      searchTerm: exampleTerm.term,
      campaignId: exampleTerm.campaignId,
      campaignName: exampleTerm.campaignName,
      addedRemoved: "--",
      adGroupId: exampleTerm.adGroupId,
      adGroupName: exampleTerm.adGroupName,
      row: exampleTerm,
      ruleLevel: ruleLevel,
      autoApply: autoApply,
      generateReport: generateReport,
      STATUS: STATUS,
      negativesToAddByAdGroup: result.negativesByAdGroup,
      negativesToAddByCampaign: result.negativesByCampaign,
      negativesToAddByAccount: result.negativesByAccount,
      outputDataRows: result.outputDataRows,
      termType: 'EXAMPLE',
      aiResponseString: "N/A"
    });

    if (autoApply) result.stats.collectedForNeg++;
  }
}

/**
 * Checks if a row meets all conditions (performance, text, and AI).
 * @param {Object} row The report row.
 * @param {string} searchTerm The search term.
 * @param {Array} performanceConditions Parsed performance conditions.
 * @param {Object} textConditions Parsed text conditions.
 * @param {Object} rule The rule object.
 * @param {string} campaignName The campaign name.
 * @param {string} adGroupName The ad group name.
 * @param {string} campaignId The campaign ID.
 * @param {string} adGroupId The ad group ID.
 * @return {boolean} True if the row meets all conditions.
 * @private
 */
function _rowMeetsAllConditions_(row, searchTerm, performanceConditions, textConditions, rule, campaignName, adGroupName, campaignId, adGroupId) {
  // Check Performance Conditions
  if (performanceConditions && !checkPerformanceConditions_(row, performanceConditions)) {
    return false;
  }

  // Check Text Conditions (true if term should be excluded)
  const shouldExcludeText = checkTextConditions_(searchTerm, textConditions);
  if (!shouldExcludeText) {
    return false;
  }

  // Check AI Prompt Condition
  const aiResponseString = getAiResponseString_(searchTerm, rule, campaignName, adGroupName, campaignId, adGroupId);
  row.aiResponseString = aiResponseString; // Store for later use

  // We'll classify "N/A" as relevant as we don't know if it's relevant or not
  const isRelevant = (response) => response.toLowerCase() === "relevant" || response.toLowerCase() === "n/a";
  const shouldExcludeAi = !isRelevant(aiResponseString);

  // If AI says it's relevant and we're not in log mode, skip
  if (!shouldExcludeAi && !rule.ai_list_all_search_terms) {
    return false;
  }

  return true;
}

/**
 * Checks if a PMAX row meets text and AI conditions.
 * @param {string} searchTerm The search term.
 * @param {Object} textConditions Parsed text conditions.
 * @param {Object} rule The rule object.
 * @param {string} campaignName The campaign name.
 * @param {string} adGroupName The ad group name (null for PMAX).
 * @param {string} campaignId The campaign ID.
 * @param {string} adGroupId The ad group ID (null for PMAX).
 * @return {boolean} True if the row meets all conditions.
 * @private
 */
function _pmaxRowMeetsConditions_(searchTerm, textConditions, rule, campaignName, adGroupName, campaignId, adGroupId) {
  // Check Text Conditions (true if term should be excluded)
  const shouldExcludeText = checkTextConditions_(searchTerm, textConditions);
  if (!shouldExcludeText) {
    return false;
  }

  // Check AI Prompt Condition
  const aiResponseString = getAiResponseString_(searchTerm, rule, campaignName, adGroupName, campaignId, adGroupId);

  // We'll classify "N/A" as relevant as we don't know if it's relevant or not
  const isRelevant = (response) => response.toLowerCase() === "relevant" || response.toLowerCase() === "n/a";
  const shouldExcludeAi = !isRelevant(aiResponseString);

  // Both text and AI conditions must indicate exclusion
  return shouldExcludeText && shouldExcludeAi;
}

/**
 * Process a matched search term (for both regular and PMAX terms).
 * @param {Object} params The parameter object containing:
 * @param {string} params.searchTerm The search term.
 * @param {string} params.campaignId The campaign ID.
 * @param {string} params.campaignName The campaign name.
 * @param {string|null} params.adGroupId The ad group ID (null for PMAX).
 * @param {string|null} params.adGroupName The ad group name (null for PMAX).
 * @param {string} params.addedRemoved The status of the search term.
 * @param {Object} params.row The data row.
 * @param {string} params.ruleLevel 'campaign' or 'adgroup'.
 * @param {boolean} params.autoApply Whether to collect negatives for auto-application.
 * @param {boolean} params.generateReport Whether to collect data for the report sheet.
 * @param {Object} params.STATUS Status message constants.
 * @param {Object} params.negativesToAddByAdGroup Negative keywords to add by ad group.
 * @param {Object} params.negativesToAddByCampaign Negative keywords to add by campaign.
 * @param {Array} params.outputDataRows Output data rows to append to.
 * @param {string} params.termType 'STANDARD' or 'PMAX' indicating the term type.
 * @param {string} params.aiResponseString The AI response for this search term.
 * @private
 */
function processMatchedTerm_(params) {
  console.log(`Processing example term ${params.searchTerm}`);
  try {
    // Determine negative status and collect negatives if auto-apply is enabled
    const negativeStatus = determineNegativeStatus_({
      termType: params.termType,
      autoApply: params.autoApply,
      STATUS: params.STATUS
    });
    console.log(`Negative status: ${negativeStatus}`);

    // Collect negatives for auto-apply if applicable - only standard and example terms can be negated
    if (params.autoApply && (params.termType === 'STANDARD' || params.termType === 'EXAMPLE' || params.termType === 'PMAX')) {
      collectNegativeKeywords_({
        searchTerm: params.searchTerm,
        campaignId: params.campaignId,
        adGroupId: params.adGroupId,
        ruleLevel: params.ruleLevel,
        negativesToAddByAdGroup: params.negativesToAddByAdGroup,
        negativesToAddByCampaign: params.negativesToAddByCampaign,
        negativesToAddByAccount: params.negativesToAddByAccount
      });
    }

    // Generate report data if needed
    if (params.generateReport) {
      generateReportData_({
        searchTerm: params.searchTerm,
        campaignName: params.campaignName,
        adGroupName: params.adGroupName,
        addedRemoved: params.addedRemoved,
        row: params.row,
        ruleLevel: params.ruleLevel,
        termType: params.termType,
        negativeStatus: negativeStatus,
        outputDataRows: params.outputDataRows,
        aiResponseString: params.aiResponseString || "N/A"
      });
    }
  } catch (e) {
    console.log(`Error in processMatchedTerm_ for term "${params.searchTerm}": ${e.message}`);
    console.log(`Stack trace: ${e.stack}`);
  }
}

/**
 * Determines the negative keyword status based on term type and application settings.
 * @param {Object} params The parameter object
 * @param {string} params.termType 'STANDARD' or 'PMAX'
 * @param {boolean} params.autoApply Whether auto-apply is enabled
 * @param {Object} params.STATUS Status message constants
 * @return {string} The negative status
 * @private
 */
function determineNegativeStatus_(params) {
  if (params.termType === 'PMAX') {
    return params.STATUS.PMAX_NOT_SUPPORTED;
  }

  if (!params.autoApply) {
    return params.STATUS.AUTO_APPLY_DISABLED;
  }

  const isPreview = AdsApp.getExecutionInfo().isPreview();
  const isTestMode = TEST_MODE && !isPreview;

  if (isTestMode) {
    return params.STATUS.TEST_MODE;
  } else if (isPreview) {
    return params.STATUS.PREVIEWED;
  } else {
    return params.STATUS.ADDED;
  }
}

/**
 * Collects negative keywords for auto-application based on rule level.
 * @param {Object} params The parameter object
 * @param {string} params.searchTerm The search term to add as negative
 * @param {string} params.campaignId The campaign ID
 * @param {string} params.adGroupId The ad group ID
 * @param {string} params.ruleLevel 'campaign' or 'adgroup'
 * @param {Object} params.negativesToAddByAdGroup Collection of ad group negatives
 * @param {Object} params.negativesToAddByCampaign Collection of campaign negatives
 * @param {Object} params.negativesToAddByAccount Collection of account negatives
 * @private
 */
function collectNegativeKeywords_(params) {
  if (params.ruleLevel === 'adgroup' && params.adGroupId) {
    addToAdGroupNegatives_({
      searchTerm: params.searchTerm,
      campaignId: params.campaignId,
      adGroupId: params.adGroupId,
      negativesToAddByAdGroup: params.negativesToAddByAdGroup
    });
  } else if (params.ruleLevel === 'campaign') {
    addToCampaignNegatives_({
      searchTerm: params.searchTerm,
      campaignId: params.campaignId,
      negativesToAddByCampaign: params.negativesToAddByCampaign
    });
  } else if (params.ruleLevel === 'account') {
    addToAccountNegatives_({
      searchTerm: params.searchTerm,
      negativesToAddByAccount: params.negativesToAddByAccount
    });
  }
}


/**
 * Adds a search term to the account negatives collection.
 * @param {Object} params The parameter object
 * @param {string} params.searchTerm The search term to add
 * @param {Object} params.negativesToAddByAccount Collection of account negatives
 * @private
 */
function addToAccountNegatives_(params) {
  if (!params.negativesToAddByAccount.terms) {
    params.negativesToAddByAccount.terms = [];
  }
  if (!params.negativesToAddByAccount.terms.includes(params.searchTerm)) {
    params.negativesToAddByAccount.terms.push(params.searchTerm);
  }
}


/**
 * Adds a search term to the ad group negatives collection.
 * @param {Object} params The parameter object
 * @param {string} params.searchTerm The search term to add
 * @param {string} params.campaignId The campaign ID
 * @param {string} params.adGroupId The ad group ID
 * @param {Object} params.negativesToAddByAdGroup Collection of ad group negatives
 * @private
 */
function addToAdGroupNegatives_(params) {
  if (!params.negativesToAddByAdGroup[params.adGroupId]) {
    params.negativesToAddByAdGroup[params.adGroupId] = { campaignId: params.campaignId, terms: [] };
  }
  if (!params.negativesToAddByAdGroup[params.adGroupId].terms.includes(params.searchTerm)) {
    params.negativesToAddByAdGroup[params.adGroupId].terms.push(params.searchTerm);
  }
}

/**
 * Adds a search term to the campaign negatives collection.
 * @param {Object} params The parameter object
 * @param {string} params.searchTerm The search term to add
 * @param {string} params.campaignId The campaign ID
 * @param {Object} params.negativesToAddByCampaign Collection of campaign negatives
 * @private
 */
function addToCampaignNegatives_(params) {
  if (!params.negativesToAddByCampaign[params.campaignId]) {
    params.negativesToAddByCampaign[params.campaignId] = { terms: [] };
  }
  if (!params.negativesToAddByCampaign[params.campaignId].terms.includes(params.searchTerm)) {
    params.negativesToAddByCampaign[params.campaignId].terms.push(params.searchTerm);
  }
}

/**
 * Generates report data for the matched term.
 * @param {Object} params The parameter object
 * @param {string} params.searchTerm The search term
 * @param {string} params.campaignName The campaign name
 * @param {string} params.adGroupName The ad group name
 * @param {string} params.addedRemoved The status of the search term
 * @param {Object} params.row The report row with metrics
 * @param {string} params.ruleLevel 'campaign' or 'adgroup'
 * @param {string} params.termType 'STANDARD' or 'PMAX'
 * @param {string} params.negativeStatus The negative keyword status
 * @param {Array} params.outputDataRows The collection to add the output row to
 * @param {string} params.aiResponseString The AI response for this search term
 * @private
 */
function generateReportData_(params) {
  try {
    const metrics = getRowMetrics_({
      row: params.row,
      termType: params.termType
    });

    if (!metrics) {
      console.log(`Warning: Could not get metrics for matched row "${params.searchTerm}". Skipping report row.`);
      return;
    }

    const outputRow = buildOutputRow_({
      searchTerm: params.searchTerm,
      campaignName: params.campaignName,
      adGroupName: params.adGroupName,
      addedRemoved: params.addedRemoved,
      metrics: metrics,
      negativeStatus: params.negativeStatus,
      ruleLevel: params.ruleLevel,
      termType: params.termType,
      aiResponseString: params.aiResponseString || "N/A"
    });

    params.outputDataRows.push(outputRow);
  } catch (metricsError) {
    console.log(`Error processing metrics for term "${params.searchTerm}": ${metricsError.message}`);
  }
}

/**
 * Extracts metrics from a report row. These need to be the same order as the header
 * @param {Object} params The parameter object
 * @param {Object} params.row The report row
 * @param {string} params.termType 'STANDARD' or 'PMAX'
 * @return {Object|null} The extracted metrics or null if extraction failed
 * @private
 */
function getRowMetrics_(params) {

  if (params.termType === 'PMAX') {
    // Use metrics directly from PMAX row
    return {
      impressions: params.row.impressions || 0,
      clicks: params.row.clicks || 0,
      cost: params.row.cost || 0,
      conversions: params.row.conversions || 0,
      conversions_value: params.row.conversions_value || 0,
      ctr: params.row.clicks > 0 && params.row.impressions > 0 ? params.row.clicks / params.row.impressions : 0,
      conversion_rate: params.row.clicks > 0 ? params.row.conversions / params.row.clicks : 0,
      cost_per_conversion: params.row.conversions > 0 ? params.row.cost / params.row.conversions : (params.row.cost > 0 ? Infinity : 0),
      cpc: params.row.clicks > 0 ? params.row.cost / params.row.clicks : (params.row.cost > 0 ? Infinity : 0),
      roas: params.row.cost > 0 ? params.row.conversions_value / params.row.cost : 0,
    };
  } else if (params.termType === 'EXAMPLE') {
    return {
      impressions: 0,
      clicks: 0,
      cost: 0,
      conversions: 0,
      conversions_value: 0,
      ctr: 0,
      conversion_rate: 0,
      cost_per_conversion: 0,
      cpc: 0,
      roas: 0,
    };
  } else {
    // Extract metrics from regular row
    return _getMetricsFromRow_(params.row);
  }
}

/**
 * Builds an output row for the report.
 * @param {Object} params The parameter object
 * @param {string} params.searchTerm The search term
 * @param {string} params.campaignName The campaign name
 * @param {string} params.adGroupName The ad group name
 * @param {Object} params.metrics The metrics object
 * @param {string} params.negativeStatus The negative status
 * @param {string} params.ruleLevel 'campaign' or 'adgroup'
 * @param {string} params.termType 'STANDARD' or 'PMAX'
 * @param {string} params.aiResponseString The AI response for this search term
 * @return {Array} The formatted output row
 * @private
 */
function buildOutputRow_(params) {
  // Ensure all metrics have default values
  const metrics = {
    impressions: 0,
    clicks: 0,
    cost: 0,
    conversions: 0,
    conversions_value: 0,
    ctr: 0,
    conversion_rate: 0,
    cost_per_conversion: 0,
    cpc: 0,
    roas: 0,
    ...params.metrics // Merge with provided metrics, overriding defaults
  };

  let outputRow = [
    params.searchTerm, params.addedRemoved,
    metrics.impressions, metrics.clicks, metrics.cost.toFixed(2),
    metrics.conversions,
    metrics.conversions_value || 0, // Ensure we display 0 if undefined
    (metrics.ctr * 100).toFixed(2) + '%',
    (metrics.conversion_rate * 100).toFixed(2) + '%',
    (metrics.cost_per_conversion === Infinity ? 'Infinity' : metrics.cost_per_conversion.toFixed(2)),
    metrics.cpc.toFixed(2),
    metrics.roas.toFixed(2),
    params.negativeStatus,
    params.aiResponseString || "N/A" // Use aiResponseString or default to N/A
  ];

  if (params.ruleLevel === 'campaign' || params.ruleLevel === 'adgroup') {
    outputRow.splice(1, 0, params.campaignName);
  }

  if (params.ruleLevel === 'adgroup' && (params.termType === 'STANDARD' || params.termType === 'EXAMPLE')) {
    // Only add ad group column for standard (non-PMAX) terms
    outputRow.splice(2, 0, params.adGroupName);
  }

  return outputRow;
}

/**
 * Get PMAX search terms for the provided campaigns.
 * @param {Array<Object>} campaigns Array of campaign objects with id and name.
 * @return {Array<Object>} Array of PMAX search term objects.
 */
function getPmaxSearchTerms(campaigns) {
  if (!campaigns || !campaigns.length) {
    return [];
  }

  const allSearchTerms = [];

  // Process each campaign one at a time
  for (const campaign of campaigns) {
    try {
      const campaignId = campaign.id;
      const campaignName = campaign.name;
      const dateCondition = _getDateRangeCondition_(this.lookbackDays);

      // Build query to get PMAX search terms for this campaign
      // Note: campaign_search_term_insight only supports specific metrics
      const query = `
                SELECT
                    campaign_search_term_insight.category_label,
                    metrics.clicks,
                    metrics.impressions,
                    metrics.conversions,
                    metrics.conversions_value
                FROM campaign_search_term_insight
                WHERE ${dateCondition}
                AND campaign_search_term_insight.campaign_id = '${campaignId}'
            `;

      console.log(`Executing PMAX search terms query for campaign ${campaignId} (${campaignName})`);

      try {
        const report = AdsApp.report(query);
        const rows = report.rows();

        let termCount = 0;
        while (rows.hasNext()) {
          const row = rows.next();
          termCount++;

          // Skip empty search terms
          if (row['campaign_search_term_insight.category_label'].trim() === '') {
            continue;
          }

          // Log if conversions_value is missing
          if (typeof row['metrics.conversions_value'] === 'undefined') {
            console.log(`Warning: PMAX search term "${row['campaign_search_term_insight.category_label']}" is missing conversions_value. Using default of 0.`);
          }

          // Format data in a way similar to regular search term report
          // For PMAX, we have to estimate cost since cost_micros isn't available
          // We'll use 0 for cost since it's not available directly
          allSearchTerms.push({
            searchTerm: row['campaign_search_term_insight.category_label'],
            campaignId: campaignId,
            campaignName: campaignName,
            impressions: parseInt(row['metrics.impressions'] || 0),
            clicks: parseInt(row['metrics.clicks'] || 0),
            cost: 0, // Cost not available for PMAX search terms
            conversions: parseFloat(row['metrics.conversions'] || 0),
            conversions_value: parseFloat(row['metrics.conversions_value'] || 0),
            isPmax: true
          });
        }

        console.log(`Retrieved ${termCount} search terms for PMAX campaign ${campaignId} (${campaignName})`);
      } catch (reportError) {
        console.log(`Error executing report for PMAX campaign ${campaignId}: ${reportError.message}`);
        console.log(`Query: ${query}`);
      }

    } catch (e) {
      console.log(`Error getting search terms for PMAX campaign ${campaign.id}: ${e.message}`);
      console.log(`Stack trace: ${e.stack}`);
      // Continue with the next campaign
    }
  }

  return allSearchTerms;
}

/**
 * Applies collected negative keywords based on the rule level and mode (Live/Preview).
 * @param {string} ruleName For logging.
 * @param {string} ruleLevel 'campaign' or 'adgroup'.
 * @param {Object} negativesByAdGroup Collected ad group negatives.
 * @param {Object} negativesByCampaign Collected campaign negatives.
 * @param {Object} negativesByAccount Collected account negatives.
 * @param {string} negativeListName The name of the negative list to apply.
 */
function _applyNegativeKeywords_(ruleName, ruleLevel, negativesByAdGroup, negativesByCampaign, negativesByAccount, negativeListName) {
  const isPreview = AdsApp.getExecutionInfo().isPreview();
  const mode = isPreview ? "PREVIEW" : "LIVE";
  let totalKeywordsSubmitted = 0;

  if (ruleLevel === 'adgroup' && !negativeListName) {
    console.log(`Applying negative keywords for rule "${ruleName}" (Ad Group Level) in ${mode} mode...`);
    for (const adGroupId in negativesByAdGroup) {
      const info = negativesByAdGroup[adGroupId];
      const termsToAdd = info.terms;
      if (termsToAdd && termsToAdd.length > 0) {
        addAdGroupNegativeKeywords(info.campaignId, adGroupId, termsToAdd);
        totalKeywordsSubmitted += termsToAdd.length;
      }
    }
    console.log(`Finished submitting negatives for rule "${ruleName}". Submitted ${totalKeywordsSubmitted} keywords across relevant ad groups.`);

  }


  if (ruleLevel === 'adgroup' && negativeListName) {
    console.log(`Applying negative keywords for rule "${ruleName}" (Ad Group Level) in ${mode} mode...`);
    let listNegativesToAdd = [];
    for (const adGroupId in negativesByAdGroup) {
      const info = negativesByAdGroup[adGroupId];
      const termsToAdd = info.terms;
      listNegativesToAdd = [...listNegativesToAdd, ...termsToAdd];
    }
    addListNegativeKeywords(listNegativesToAdd, negativeListName);
    console.log(`Finished submitting negatives for rule "${ruleName}". Submitted ${totalKeywordsSubmitted} keywords across relevant campaigns.`);
  }

  if (ruleLevel === 'campaign' && !negativeListName) {
    console.log(`Applying negative keywords for rule "${ruleName}" (Campaign Level) in ${mode} mode...`);
    for (const campaignId in negativesByCampaign) {
      const info = negativesByCampaign[campaignId];
      const termsToAdd = info.terms;
      if (termsToAdd && termsToAdd.length > 0) {
        addCampaignNegativeKeywords(campaignId, termsToAdd);
        totalKeywordsSubmitted += termsToAdd.length;
      }
    }
    console.log(`Finished submitting negatives for rule "${ruleName}". Submitted ${totalKeywordsSubmitted} keywords across relevant campaigns.`);

  }

  if (ruleLevel === 'campaign' && negativeListName) {
    console.log(`Applying negative keywords for rule "${ruleName}" (Campaign Level) in ${mode} mode...`);
    let listNegativesToAdd = [];
    for (const campaignId in negativesByCampaign) {
      const info = negativesByCampaign[campaignId];
      const termsToAdd = info.terms;
      listNegativesToAdd = [...listNegativesToAdd, ...termsToAdd];
    }
    addListNegativeKeywords(listNegativesToAdd, negativeListName);
    console.log(`Finished submitting negatives for rule "${ruleName}". Submitted ${totalKeywordsSubmitted} keywords across relevant campaigns.`);
  }

  if (ruleLevel === 'account') {
    if (!negativeListName) {
      console.error(`No negative list name provided. Skipping application of negatives.`);
      return;
    }
    addListNegativeKeywords(negativesByAccount.terms, negativeListName);
    return; // Exit early if level isn't handled
  }

  // Add preview mode note if applicable
  if (isPreview) {
    console.log("NOTE: In Preview Mode, changes are logged but not permanently applied.");
  }
}

/**
 * Adds a list of negative keywords to a specific list.
 * @param {Array<string>} termsToAdd An array of search terms to add as negatives.
 * @param {string} negativeListName The name of the negative list to add to.
 */
function addListNegativeKeywords(termsToAdd, negativeListName) {
  if (!termsToAdd || termsToAdd.length === 0) {
    return;
  }
  console.log(`Adding ${termsToAdd.length} negative keywords to list "${negativeListName}"`);
  //get the list
  let listIter = AdsApp.negativeKeywordLists().withCondition(`shared_set.name = "${negativeListName}"`).get();
  //console error if not exists
  if (!listIter.hasNext()) {
    console.error(`ERROR: Negative keyword list "${negativeListName}" not found.`);
    return;
  }
  let list = listIter.next();
  //add match types
  const negativeKeywords = termsToAdd.map(term => addMatchType(term, NEGATIVE_MATCH_TYPE));
  let first5 = negativeKeywords.slice(0, 5);
  console.log("\nFirst 5 negatives preview: ");
  first5.forEach(term => console.log(`Adding ${term} to list "${negativeListName}"`));
  console.log("\n");
  //add terms to list
  list.addNegativeKeywords(negativeKeywords);
}


/**
 * Adds a list of negative keywords to a specific ad group.
 * Handles checking for both standard Search and Shopping ad group types.
 * @param {string} campaignId The ID of the campaign containing the ad group.
 * @param {string} adGroupId The ID of the ad group to add negatives to.
 * @param {Array<string>} negativeKeywords An array of search terms to add as negatives.
 * @param {string} [campaignType=CAMPAIGN_TYPES.search] The type of campaign ('Search' or 'Shopping').
 * @private // Assuming this convention for helper functions
 */
function addAdGroupNegativeKeywords(campaignId, adGroupId, negativeKeywords, campaignType = CAMPAIGN_TYPES.search) {
  console.log(`\nAdding ${negativeKeywords.length} exact match negatives to Ad Group ID: ${adGroupId} in Campaign ID: ${campaignId}`);
  if (!campaignId || !adGroupId || !negativeKeywords || negativeKeywords.length === 0) {
    console.log("Skipping addAdGroupNegativeKeywords due to missing IDs or empty keyword list.");
    return;
  }

  // We need to handle both Search and Shopping campaigns potentially
  let adGroupIterator;
  if (campaignType === CAMPAIGN_TYPES.shopping) {
    adGroupIterator = AdsApp.shoppingAdGroups()
      .withIds([adGroupId])
      .withCondition(`campaign.id = '${campaignId}'`)
      .get();
  } else {
    adGroupIterator = AdsApp.adGroups()
      .withIds([adGroupId])
      .withCondition(`campaign.id = '${campaignId}'`)
      .get();
  }


  if (!adGroupIterator.hasNext() && campaignType === CAMPAIGN_TYPES.search) {
    //this will happen if it's a shopping campaign
    addAdGroupNegativeKeywords(campaignId, adGroupId, negativeKeywords, CAMPAIGN_TYPES.shopping);
    return;
  }

  if (!adGroupIterator.hasNext() && campaignType === CAMPAIGN_TYPES.shopping) {
    // If still not found after checking both types
    console.log(`ERROR: Ad Group with ID ${adGroupId} not found within Campaign ID ${campaignId} (checked Search and Shopping). Cannot add negatives.`);
    return; // Stop if the ad group isn't found
  }

  // Should only be one ad group with this ID
  const adGroup = adGroupIterator.next();
  adGroupNegativeKeywordService(adGroup, negativeKeywords);

}


/**
 * Service function to add multiple negative keywords to a single AdGroup entity.
 * @param {AdsApp.AdGroup | AdsApp.ShoppingAdGroup} adGroup The AdGroup or ShoppingAdGroup entity.
 * @param {Array<string>} negativeKeywords The array of keyword texts to add.
 */
function adGroupNegativeKeywordService(adGroup, negativeKeywords) {
  console.log(`Adding ${negativeKeywords.length} negatives to Ad Group ID: ${adGroup.getId()}`);
  for (const keywordText of negativeKeywords) {
    // Check if keywordText is valid before formatting
    if (!keywordText || typeof keywordText !== 'string' || keywordText.trim() === '') {
      console.log(`Warning: Skipping invalid keyword text: "${keywordText}"`);
      continue;
    }
    const negativeKeywordFormatted = addMatchType(keywordText, NEGATIVE_MATCH_TYPE);
    // Check if formatting returned null (e.g., empty string after trim)
    if (!negativeKeywordFormatted) {
      console.log(`Warning: Skipping keyword "${keywordText}" due to formatting issue.`);
      continue;
    }
    try {
      adGroup.createNegativeKeyword(negativeKeywordFormatted);
    } catch (e) {
      // Log errors but continue trying to add others
      console.log(`  ERROR adding negative "${negativeKeywordFormatted}" to Ad Group ID ${adGroup.getId()}: ${e}`);
    }
  }
}

/**
 * Adds a list of negative keywords to a specific campaign.
 * Handles checking for both standard Search and Shopping campaign types.
 * @param {string} campaignId The ID of the campaign to add negatives to.
 * @param {Array<string>} negativeKeywords An array of search terms to add as negatives.
 * @param {string} [campaignType=CAMPAIGN_TYPES.search] The type of campaign ('Search' or 'Shopping').
 * @private
 */
function addCampaignNegativeKeywords(campaignId, negativeKeywords, campaignType = CAMPAIGN_TYPES.search) {
  console.log(`\nAdding ${negativeKeywords.length} exact match negatives to Campaign ID: ${campaignId}`);
  if (!campaignId || !negativeKeywords || negativeKeywords.length === 0) {
    console.log("Skipping addCampaignNegativeKeywords due to missing ID or empty keyword list.");
    return;
  }

  // We need to handle both Search and Shopping campaigns potentially
  let campaignIterator;
  if (campaignType === CAMPAIGN_TYPES.shopping) {
    campaignIterator = AdsApp.shoppingCampaigns()
      .withIds([campaignId])
      .get();
  } else {
    campaignIterator = AdsApp.campaigns()
      .withIds([campaignId])
      .get();
  }

  if (!campaignIterator.hasNext() && campaignType === CAMPAIGN_TYPES.search) {
    //this will happen if it's a shopping campaign
    addCampaignNegativeKeywords(campaignId, negativeKeywords, CAMPAIGN_TYPES.shopping);
    return;
  }

  if (!campaignIterator.hasNext() && campaignType === CAMPAIGN_TYPES.shopping) {
    // If still not found after checking both types
    console.log(`ERROR: Campaign with ID ${campaignId} not found (checked Search and Shopping). Cannot add negatives.`);
    return; // Stop if the campaign isn't found
  }

  // Should only be one campaign with this ID
  const campaign = campaignIterator.next();
  campaignNegativeKeywordService(campaign, negativeKeywords);
}

/**
 * Service function to add multiple negative keywords to a single Campaign entity.
 * @param {AdsApp.Campaign | AdsApp.ShoppingCampaign} campaign The Campaign or ShoppingCampaign entity.
 * @param {Array<string>} negativeKeywords The array of keyword texts to add.
 */
function campaignNegativeKeywordService(campaign, negativeKeywords) {
  console.log(`Adding ${negativeKeywords.length} negatives to Campaign ID: ${campaign.getId()}`);
  for (const keywordText of negativeKeywords) {
    // Check if keywordText is valid before formatting
    if (!keywordText || typeof keywordText !== 'string' || keywordText.trim() === '') {
      console.log(`Warning: Skipping invalid keyword text: "${keywordText}"`);
      continue;
    }
    const negativeKeywordFormatted = addMatchType(keywordText, NEGATIVE_MATCH_TYPE);
    // Check if formatting returned null (e.g., empty string after trim)
    if (!negativeKeywordFormatted) {
      console.log(`Warning: Skipping keyword "${keywordText}" due to formatting issue.`);
      continue;
    }
    try {
      campaign.createNegativeKeyword(negativeKeywordFormatted);
    } catch (e) {
      // Log errors but continue trying to add others
      console.log(`  ERROR adding negative "${negativeKeywordFormatted}" to Campaign ID ${campaign.getId()}: ${e}`);
    }
  }
}

/**
 * Formats a keyword string with the specified match type syntax.
 * @param {string} word The keyword text.
 * @param {string} matchType The desired match type ('EXACT', 'PHRASE', 'BROAD', 'BMM').
 * @return {string} The formatted keyword string.
 * @throws {Error} If the match type is not recognized.
 */
function addMatchType(word, matchType) {
  word = String(word).trim(); // Ensure it's a string and trim whitespace
  // Handle potential empty strings after trimming
  if (!word) {
    console.log("Warning: Attempted to format an empty string as a negative keyword.");
    return null; // Return null or handle as appropriate
  }
  const matchTypeLower = matchType.toLowerCase();

  if (matchTypeLower == "broad") {
    return word; // No modification for broad
  } else if (matchTypeLower == "bmm") {
    // BMM is deprecated, but keeping logic in case needed elsewhere.
    // For negatives, Broad is usually preferred over BMM now.
    return word.split(/\s+/).map(x => `+${x}`).join(" ");
  } else if (matchTypeLower == "phrase") {
    // Add quotes if not already present
    return word.startsWith('"') && word.endsWith('"') ? word : `"${word}"`;
  } else if (matchTypeLower == "exact") {
    // Add brackets if not already present
    return word.startsWith('[') && word.endsWith(']') ? word : `[${word}]`;
  } else {
    throw new Error(`Match type "${matchType}" not recognised. Please provide one of Broad, Phrase, or Exact.`);
  }
}

/**
 * Builds the GAQL query string based on rule parameters.
 * @param {string} level 'campaign' or 'adgroup'.
 * @param {number} lookbackDays Number of days to look back.
 * @param {Array<Object>} entityConditions Parsed entity conditions.
 * @param {number|null} maxSearchTerms Optional maximum number of search terms to return.
 * @param {Array<Object>} performanceConditions Parsed performance conditions.
 * @param {Object} searchTermStatusParams Search term status parameters (searchTermStatusAdded, searchTermStatusExcluded, searchTermStatusAddedExcluded)
 * @return {string} The GAQL query string.
 * @throws {Error} If lookbackDays is invalid or entity conditions format is invalid.
 * @private
 */
function buildGaqlQuery_(level, lookbackDays, entityConditions, maxSearchTerms = null, performanceConditions = null, searchTermStatusParams) {
  if (isNaN(parseInt(lookbackDays)) || lookbackDays < 0) {
    throw new Error("Invalid lookback_days value: " + lookbackDays);
  }

  // Define core fields needed always
  const baseFields = [
    'search_term_view.search_term',
    'search_term_view.status',
    // Ad group fields added conditionally below
    'metrics.impressions', 'metrics.clicks',
    'metrics.cost_micros', 'metrics.conversions', 'metrics.conversions_value'
  ];

  // Conditionally add campaign fields if level is 'adgroup' or 'campaign'
  let fields = [...baseFields]; // Start with a copy
  if (level && (level.toLowerCase() === 'adgroup' || level.toLowerCase() === 'campaign')) {
    fields.splice(2, 0, 'campaign.id', 'campaign.name'); // Insert campaign fields after search term fields
  }

  // Conditionally add Ad Group fields if level is 'adgroup'
  fields = [...fields]; // Start with a copy
  if (level && level.toLowerCase() === 'adgroup') {
    fields.splice(3, 0, 'ad_group.id', 'ad_group.name'); // Insert ad group fields after campaign fields
  }

  const selectClause = 'SELECT ' + fields.join(', ');
  const fromClause = 'FROM search_term_view';

  // --- Build WHERE Clause ---
  let whereConditions = [];

  // 1. Date Range Filter
  whereConditions.push(_getDateRangeCondition_(lookbackDays));

  // 2. Status Filters
  whereConditions = whereConditions.concat(_getStatusConditions_());
  // 2.1 Added/Excluded (Searc Term Status) Filters
  if (Object.values(searchTermStatusParams).some(value => value === false)) {
    whereConditions = whereConditions.concat(_getSearchTermStatusConditions_(searchTermStatusParams));
  }

  // 3. Entity Condition Filters
  if (!Array.isArray(entityConditions)) {
    throw new Error("Invalid entity_conditions format: Expected an array. Received: " + JSON.stringify(entityConditions));
  }
  entityConditions.forEach(condition => {
    // Skip ad_group conditions if the level is 'campaign'
    if (level && level.toLowerCase() !== 'adgroup' && condition.field && condition.field.startsWith('ad_group.')) {
      return;
    }
    const conditionString = _getEntityConditionString_(condition);
    if (conditionString) { // Only add valid, non-null conditions
      whereConditions.push(conditionString);
    }
  });

  // 4. Performance Condition Filters
  whereConditions = _addPerformanceConditionsToQuery(performanceConditions, whereConditions);

  // --- Assemble Query ---
  const whereClause = 'WHERE ' + whereConditions.join(' AND ');
  let gaqlQuery = `${selectClause} ${fromClause} ${whereClause}`;

  // Add LIMIT clause if maxSearchTerms is specified
  if (maxSearchTerms && !isNaN(maxSearchTerms) && maxSearchTerms > 0) {
    gaqlQuery += ` LIMIT ${maxSearchTerms}`;
    console.log(`Applied search terms limit of ${maxSearchTerms} to GAQL query.`);
  }

  console.log(`Built GAQL Query (Level: ${level || 'undefined'}): ${gaqlQuery}`);
  return gaqlQuery;
}

function _getSearchTermStatusConditions_(searchTermStatusParams) {
  const { searchTermStatusAddedExcluded, searchTermStatusAdded, searchTermStatusExcluded } = searchTermStatusParams;
  let statuses = ['NONE', 'UNKNOWN'];
  if (searchTermStatusAddedExcluded) {
    statuses.push('ADDED_EXCLUDED');
  }
  if (searchTermStatusAdded) {
    statuses.push('ADDED');
  }
  if (searchTermStatusExcluded) {
    statuses.push('EXCLUDED');
  }
  // Convert statuses array to a string with single quotes
  const statusesString = statuses.map(status => `'${status}'`).join(',');
  const statusCondition = `search_term_view.status IN (${statusesString})`;
  return statusCondition;
}



/**
 * Checks if a report row meets the performance conditions.
 * @param {Object} row A report row object from AdsApp.report().rows().next().
 * @param {Array<Object>} performanceConditions Parsed performance conditions.
 * @return {boolean} True if all conditions pass, false otherwise.
 * @private
 */
function checkPerformanceConditions_(row, performanceConditions) {
  if (!performanceConditions || performanceConditions.length === 0) {
    return true; // No conditions to check, so it passes.
  }
  if (!Array.isArray(performanceConditions)) {
    console.log("Error: performanceConditions is not an array. Skipping checks.");
    return false; // Treat malformed conditions as failure
  }

  const metrics = _getMetricsFromRow_(row);
  if (!metrics) {
    console.log("Warning: Could not extract metrics from row. Performance check fails.");
    return false; // Cannot check conditions if metrics are unavailable
  }

  // Uncomment for detailed debugging of metrics per row
  // console.log("Checking Perf Conditions for Row. Metrics: " + JSON.stringify(metrics));

  for (const condition of performanceConditions) {
    const metricName = condition.metric;
    const operator = condition.operator;
    // Ensure the threshold value is parsed as a number
    const thresholdValue = parseFloat(condition.value);

    if (!metricName || !operator || isNaN(thresholdValue)) {
      // console.log(`Warning: Skipping invalid performance condition format: ${JSON.stringify(condition)}`);
      continue; // Skip malformed conditions
    }

    // Skip base metrics that are already filtered in the GAQL query
    // unless we're in PMAX processing where the GAQL filter wasn't applied
    const isBaseMetric = Object.keys(BASE_METRICS).some(key => key.split('.')[1] === metricName);
    if (isBaseMetric && row.query_source !== 'PMAX') {
      // console.log(`Skipping base metric "${metricName}" check as it was already filtered in GAQL query`);
      continue;
    }

    if (!(metricName in metrics)) {
      // console.log(`Warning: Metric "${metricName}" not found in calculated metrics. Skipping condition: ${JSON.stringify(condition)}`);
      continue; // Skip if the required metric wasn't calculated
    }

    const metricValue = metrics[metricName];

    try {
      const conditionMet = _compareMetric_(metricValue, operator, thresholdValue);
      // Uncomment for detailed debugging of each comparison
      // console.log(`-- Checking: ${metricName} (${metricValue}) ${operator} ${thresholdValue} -> ${conditionMet}`);
      if (!conditionMet) {
        // console.log(`-----> Perf Condition Failed: ${metricName} (${metricValue}) ${operator} ${thresholdValue}`);
        return false; // If any condition fails, the row fails.
      }
    } catch (e) {
      // console.log(`Error comparing performance condition: ${JSON.stringify(condition)}. Error: ${e.message}. Skipping row check.`);
      return false; // Treat errors during comparison as failure for this row
    }
  }

  // If the loop completes without returning false, all conditions passed.
  // console.log("All Perf Conditions Passed.");
  return true;
}

/**
 * Extracts base metrics and calculates derived metrics from a report row.
 * Handles potential null/undefined values and division by zero.
 * Cost is converted from micros to standard currency units.
 * @param {Object} row A report row object from AdsApp.report().rows().next().
 * @return {Object} An object containing key metrics (impressions, clicks, cost, conversions, ctr, conversion_rate, cost_per_conversion). Returns null if essential metrics are missing.
 * @private
 */
function _getMetricsFromRow_(row) {
  // Check if the row itself exists
  if (!row) {
    console.log("Warning: Received null or undefined row object.");
    return null;
  }
  // We no longer expect a nested 'metrics' object when using AdsApp.report
  // We access fields directly like row['metrics.impressions']


  // Check for essential metric fields before proceeding
  const requiredMetrics = ['metrics.impressions', 'metrics.clicks', 'metrics.cost_micros', 'metrics.conversions', 'metrics.conversions_value'];
  for (const metricKey of requiredMetrics) {
    if (typeof row[metricKey] === 'undefined') {
      // It's possible for metrics to be genuinely zero, but undefined suggests the field wasn't returned or accessed correctly.
      console.log(`Warning: Metric key "${metricKey}" not found in report row. Row data: ${JSON.stringify(row)}`);
      // Decide if this is critical. If we *always* expect these metrics, maybe return null.
      // If zero values are acceptable and represented as missing fields sometimes, provide defaults. Let's default to 0 for calculations.
      // Returning null might be safer if subsequent logic depends heavily on these. Let's return null for safety.
      return null;
    }
  }


  const metrics = {
    // Access metrics directly using flattened keys, provide default 0 if null/undefined (though check above makes this less likely needed)
    impressions: parseInt(row['metrics.impressions'] || 0),
    clicks: parseInt(row['metrics.clicks'] || 0),
    // Convert cost from micros (millionths) to standard currency units
    cost: (parseInt(row['metrics.cost_micros'] || 0)) / 1000000.0,
    conversions: parseFloat(row['metrics.conversions'] || 0), // Conversions can be fractional
    conversions_value: parseFloat(row['metrics.conversions_value'] || 0), // Make conversions_value optional with default 0
    ctr: 0,
    conversion_rate: 0,
    cost_per_conversion: 0,
    cpc: 0,
    roas: 0,
  };

  // Calculate derived metrics safely
  if (metrics.impressions > 0) {
    metrics.ctr = metrics.clicks / metrics.impressions;
  }

  if (metrics.clicks > 0) {
    metrics.conversion_rate = metrics.conversions / metrics.clicks;
  }

  if (metrics.conversions > 0) {
    metrics.cost_per_conversion = metrics.cost / metrics.conversions;
  } else if (metrics.cost > 0) {
    // Handle case where cost > 0 but conversions = 0. Assign Infinity or keep 0?
    // Assigning Infinity aligns with mathematical definition, helps '> x' checks.
    // Assigning 0 might be misleading for '< x' checks. Let's use Infinity.
    metrics.cost_per_conversion = Infinity;
  }
  // else cost is 0 and conversions are 0, cost_per_conversion remains 0.

  if (metrics.clicks > 0) {
    metrics.cpc = metrics.cost / metrics.clicks;
  }

  if (metrics.cost > 0) {
    metrics.roas = metrics.conversions_value / metrics.cost;
  }

  return metrics;
}

/**
 * Compares a metric value against a threshold using a specified operator.
 * @param {number} metricValue The actual value of the metric.
 * @param {string} operator The comparison operator ('>', '<', '=', '>=', '<=').
 * @param {number} thresholdValue The value from the performance condition.
 * @return {boolean} True if the condition is met, false otherwise.
 * @throws {Error} If the operator is invalid.
 * @private
 */
function _compareMetric_(metricValue, operator, thresholdValue) {
  switch (operator) {
    case '>':
      return metricValue > thresholdValue;
    case '<':
      return metricValue < thresholdValue;
    case '=':
      // Use a small epsilon for floating point comparison if needed, but direct equality is often expected here.
      return metricValue === thresholdValue;
    case '>=':
      return metricValue >= thresholdValue;
    case '<=':
      return metricValue <= thresholdValue;
    default:
      throw new Error(`Unsupported performance operator: "${operator}"`);
  }
}

/**
 * Checks if a search term meets the text conditions defined in the rule
 * by utilizing the global SearchTermMatcher's shouldExclude method.
 * @param {string} searchTerm The search term string.
 * @param {Object} textConditions Parsed text conditions object (e.g., {sections: [{operator: 'OR', conditions: [...]}]}).
 * @return {boolean} True if the search term FAILS the conditions (i.e., should be excluded), false otherwise.
 * @private
 */
function checkTextConditions_(searchTerm, textConditions) {
  // If no conditions/sections defined in the rule, the term should not be excluded by default.
  if (!textConditions || !Array.isArray(textConditions.sections) || textConditions.sections.length === 0) {
    // console.log(`No valid text condition sections found for term "${searchTerm}". Term should not be excluded.`);
    return true; // Should be excluded
  }

  // Assume SearchTermMatcher is globally available
  let matcher;
  try {
    matcher = new SearchTermMatcher();
  } catch (e) {
    console.log(`FATAL ERROR: Could not instantiate or access global SearchTermMatcher: ${e.message}. Text checks cannot proceed.`);
    // Fail safe: if matcher isn't available, assume the term should not be excluded.
    return false;
  }

  try {
    // shouldExclude returns true if the term FAILS the conditions (and should be excluded).
    const shouldExcludeTerm = matcher.shouldExclude(searchTerm, textConditions);
    // console.log(`SearchTermMatcher.shouldExclude("${searchTerm}", ...) returned: ${shouldExcludeTerm}.`);
    return shouldExcludeTerm; // Directly return the exclusion flag

  } catch (e) {
    // Log errors during the matching process
    console.log(`Error during SearchTermMatcher.shouldExclude for term "${searchTerm}" with conditions ${JSON.stringify(textConditions)}: ${e.message}`);
    // Fail safe: if an error occurs during matching, assume the term should not be excluded.
    return false;
  }
}

/**
 * Writes data to a specific sheet (tab) in the spreadsheet. Creates the sheet if it doesn't exist.
 * Handles sheet name length limits and writes data or a 'no results' message.
 * @param {Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @param {string} ruleName The name of the rule.
 * @param {Array<Array>} data The 2D array of data to write (assumes first row is headers).
 * @param {string|number} ruleNumber The ID of the rule to use as the sheet name.
 * @param {string|number} ruleId The ID of the rule to use as the sheet name.
 * @param {Sheet} sheet The sheet object to write to.
 * @private
 */
function writeToSheet_(spreadsheet, ruleName, data, ruleNumber, ruleId, sheet) {
  // Format the sheet name for consistency
  const formattedSheetName = isNumeric(ruleNumber) ? `${ruleNumber}` : String(ruleNumber);
  console.log(`Attempting to write data to sheet: "${formattedSheetName}" for rule "${ruleName}"`);

  try {
    _clearSheetContent_(sheet);

    // Add the rule info header section
    const spreadsheetId = extractSpreadsheetId_(SPREADSHEET_URL);
    const ruleExecutionTimestamp = Utilities.formatDate(new Date(), AdsApp.currentAccount().getTimeZone(), "yyyy-MM-dd HH:mm:ss");
    const editUrl = `https://autoneg.shabba.io/?id=${spreadsheetId}&ruleNumber=${ruleNumber}`;

    // Set up the header cells
    sheet.getRange(RULE_NAME_RANGE).setValue(ruleName);
    sheet.getRange(RULE_NAME_RANGE).setFontWeight('bold').setFontSize(14);

    sheet.getRange(RULE_LAST_RUNTIME_LABEL_RANGE).setValue('Last runtime:');
    sheet.getRange(RULE_LAST_RUNTIME_LABEL_RANGE).setFontWeight('bold');

    sheet.getRange(RULE_LAST_RUNTIME_VALUE_RANGE).setValue(ruleExecutionTimestamp);

    // Add email tracking cells
    sheet.getRange(RULE_LAST_EMAIL_LABEL_RANGE).setValue('Last email sent time:');
    sheet.getRange(RULE_LAST_EMAIL_LABEL_RANGE).setNote('Used to enforce the 1 email per day limit');
    sheet.getRange(RULE_LAST_EMAIL_LABEL_RANGE).setFontWeight('bold');
    // B3 will store the timestamp when email is sent

    // Add the "Edit rule" hyperlink
    sheet.getRange(RULE_EDIT_LINK_RANGE).setValue('Edit rule');
    sheet.getRange(RULE_EDIT_LINK_RANGE).setFontColor('blue').setFontLine('underline');
    sheet.getRange(RULE_EDIT_LINK_RANGE).setFormula(`=HYPERLINK("${editUrl}","Edit rule")`);


    sheet.setFrozenRows(RULE_DATA_START_ROW);

    // Write the data starting from row 5
    _writeDataOrMessage_(sheet, data, RULE_DATA_START_ROW);

    console.log(`Successfully finished writing to sheet: "${formattedSheetName}"`);

  } catch (e) {
    // Log detailed error
    console.log(`ERROR writing to sheet "${formattedSheetName}": ${e}`);
    if (e.message) console.log(`Error message: ${e.message}`);
    if (e.stack) console.log(`Stack trace: ${e.stack}`);
  }
}

/**
 * Gets a sheet by name, creating it if it doesn't exist.
 * @param {Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @param {string} sheetName The exact name of the sheet to get or create.
 * @return {Sheet} The sheet object.
 * @throws Error if sheet cannot be retrieved or created.
 * @private
 */
function _getOrCreateSheet_(spreadsheet, sheetName) {
  // Format the sheet name to ensure it's not treated as a number
  const formattedSheetName = isNumeric(sheetName) ? `${sheetName}` : String(sheetName);

  let sheet = spreadsheet.getSheetByName(formattedSheetName);
  if (!sheet) {
    console.log(`Sheet "${formattedSheetName}" not found. Creating it.`);
    sheet = spreadsheet.insertSheet(formattedSheetName);
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(spreadsheet.getNumSheets()); // Move to end
  } else {
    console.log(`Found existing sheet: "${sheetName}".`);
  }
  if (!sheet) { // Defensive check in case insertSheet fails silently (unlikely)
    throw new Error(`Failed to get or create sheet named "${sheetName}"`);
  }
  return sheet;
}

/**
 * Clears the content of a sheet.
 * @param {Sheet} sheet The sheet object to clear.
 * @private
 */
function _clearSheetContent_(sheet) {
  console.log(`Clearing contents of sheet: "${sheet.getName()}"`);
  // Check if there's any content to clear
  const lastRow = sheet.getLastRow();
  if (lastRow >= RULE_DATA_START_ROW) {
    const reportRange = sheet.getRange(RULE_DATA_START_ROW, 1, lastRow - RULE_DATA_START_ROW + 1, sheet.getLastColumn());
    reportRange.clearContent();
  } else {
    console.log(`No content to clear in sheet "${sheet.getName()}" as it contains less than ${RULE_DATA_START_ROW} rows.`);
  }
}

/**
 * Writes the provided 2D data array to the specified sheet.
 * Handles empty data scenarios, ensures sheet dimensions are adequate,
 * and reapplies a filter to the data range (or removes it if no data).
 * @param {Sheet} sheet The target sheet object.
 * @param {Array<Array>} data The 2D array of data (first row assumed headers).
 * @param {number} [startRow=1] The row to start writing data at (default: 1).
 * @private
 */
function _writeDataOrMessage_(sheet, data, startRow = 1) {
  // Remove any existing filter before potential structural changes or clearing
  const existingFilter = sheet.getFilter();
  if (existingFilter) {
    try {
      existingFilter.remove();
      console.log(`Removed existing filter from sheet "${sheet.getName()}" before writing.`);
    } catch (removeFilterError) {
      console.log(`Warning: Could not remove existing filter from sheet "${sheet.getName()}". Error: ${removeFilterError}`);
      // Proceed cautiously, hoping clearContents or subsequent createFilter works
    }
  }

  // Check if data is valid and has more than just a header row
  if (data && data.length > 1 && data[0] && data[0].length > 0) {
    const numRows = data.length;
    const numCols = data[0].length;
    console.log(`Writing ${numRows - 1} data rows (${numCols} columns) to sheet: "${sheet.getName()}" starting at row ${startRow}`);

    // Ensure sheet is large enough before writing
    const requiredRows = startRow + numRows - 1;
    const currentMaxRows = sheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
      sheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }
    const currentMaxCols = sheet.getMaxColumns();
    if (currentMaxCols < numCols) {
      sheet.insertColumnsAfter(currentMaxCols, numCols - currentMaxCols);
    }

    // Write data starting at the specified row
    const dataRange = sheet.getRange(startRow, 1, numRows, numCols);
    dataRange.setValues(data);

    // Re-apply the filter to the new data range (including headers)
    try {
      dataRange.createFilter();
      console.log(`Applied filter to range ${dataRange.getA1Notation()} on sheet "${sheet.getName()}".`);
    } catch (createFilterError) {
      console.log(`Warning: Could not apply filter to sheet "${sheet.getName()}". Error: ${createFilterError}`);
    }

  } else {
    // Handle case with no data or only headers
    const message = "No matching search terms found for this rule.";
    console.log(`${message} Writing message to sheet: "${sheet.getName()}" at row ${startRow}`);
    sheet.getRange(startRow, 1).setValue(message);

    // Clean up unused rows/columns if sheet might have had old data
    const lastRow = sheet.getLastRow();
    if (lastRow > startRow) {
      sheet.deleteRows(startRow + 1, lastRow - startRow);
    }
    const lastCol = sheet.getLastColumn();
    if (lastCol > 1) {
      sheet.deleteColumns(2, lastCol - 1);
    }
    // Ensure no filter exists on an empty/message-only sheet
    const finalFilterCheck = sheet.getFilter();
    if (finalFilterCheck) {
      try { finalFilterCheck.remove(); } catch (e) { /* Ignore */ }
    }
  }
}

/**
 * Extracts the Spreadsheet ID from a Google Sheet URL.
 * @param {string} url The Google Sheet URL.
 * @return {string|null} The extracted ID or null if not found.
 * @private
 */
function extractSpreadsheetId_(url) {
  const match = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

/**
 * Translates a single entity condition object into a GAQL condition string.
 * @param {Object} condition The condition object (e.g., {field: '...', operator: '...', value: '...'}).
 * @return {string|null} The GAQL condition string or null if invalid/unsupported.
 * @private
 */
function _getEntityConditionString_(condition) {
  const field = condition.field;
  const operator = condition.operator;
  let value = condition.value;

  if (!field || !operator || typeof value === 'undefined') {
    console.log("Skipping invalid entity condition format: " + JSON.stringify(condition));
    return null;
  }

  // Escape single quotes in the value for GAQL string literals
  // Ensure value is treated as a string before replacing
  const escapedValue = String(value).replace(/'/g, "\\'");

  switch (operator.toLowerCase()) {
    case 'contains':
      // Use normaliseLikeString to properly escape special characters in LIKE patterns
      return `${field} LIKE '%${normaliseLikeString(escapedValue)}%'`;
    case 'not_contains':
      // Use normaliseLikeString to properly escape special characters in LIKE patterns
      return `${field} NOT LIKE '%${normaliseLikeString(escapedValue)}%'`;
    case 'regex_contains':
      // GAQL uses RE2 syntax. Assumes input regex is valid RE2.
      return `${field} REGEXP_MATCH '${escapedValue}'`;
    case 'not_regex_contains':
      return `NOT ${field} REGEXP_MATCH '${escapedValue}'`;
    case 'equals':
    case '=':
      return `${field} = '${escapedValue}'`;
    // Add other operators like 'starts_with', 'ends_with' if needed
    // case 'starts_with':
    //     return `${field} LIKE '${normaliseLikeString(escapedValue)}%'`;
    // case 'ends_with':
    //      return `${field} LIKE '%${normaliseLikeString(escapedValue)}'`;
    default:
      console.log("Unsupported entity operator: '" + operator + "' in condition: " + JSON.stringify(condition) + ". Skipping.");
      return null;
  }
}

/**
 * Generates the GAQL date range condition string.
 * @param {number} lookbackDays Number of days to look back.
 * @return {string} The GAQL condition string (e.g., "segments.date BETWEEN 'YYYY-MM-DD' AND 'YYYY-MM-DD'").
 * @private
 */
function _getDateRangeCondition_(lookbackDays) {
  const endDate = getGoogleAdsApiFormattedDate(0); // Today
  const startDate = getGoogleAdsApiFormattedDate(lookbackDays);
  return `segments.date BETWEEN '${startDate}' AND '${endDate}'`;
}

/**
 * Returns a date string in YYYY-MM-DD format for a given number of days ago.
 * @param {number} daysAgo - Number of days to go back from today (0 = today, 1 = yesterday, etc.)
 * @returns {string} Date string in YYYY-MM-DD format
 */
function getGoogleAdsApiFormattedDate(daysAgo = 0) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);

  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');

  return `${year}-${month}-${day}`;
}


/**
 * Returns an array of standard GAQL status condition strings.
 * @return {Array<string>} Array of status conditions.
 * @private
 */
function _getStatusConditions_() {
  return [
    'campaign.status = ENABLED',
    'ad_group.status = ENABLED',
  ];
}

// 4. Performance Condition Filters
function _addPerformanceConditionsToQuery(performanceConditions, whereConditions) {
  // Return early if no valid conditions
  if (!performanceConditions || !Array.isArray(performanceConditions) || performanceConditions.length === 0) {
    return whereConditions;
  }

  // console.log("Adding performance conditions to GAQL query");
  for (const condition of performanceConditions) {
    const metricName = condition.metric;
    const operator = condition.operator;
    const thresholdValue = parseFloat(condition.value);

    // Skip if any required field is missing or invalid
    if (!metricName || !operator || isNaN(thresholdValue)) {
      // console.log(`Skipping invalid performance condition: ${JSON.stringify(condition)}`);
      continue;
    }

    // Skip non-base metrics
    if (!_isBaseMetric(metricName)) {
      // console.log(`Skipping non-base metric "${metricName}" in GAQL - will be filtered post-query`);
      continue;
    }

    // Process valid base metric
    const { metricField, metricValue } = _formatMetricForQuery(metricName, thresholdValue);
    const gaqlOperator = operator === '=' ? '=' : operator;

    whereConditions.push(`${metricField} ${gaqlOperator} ${metricValue}`);
    // console.log(`Added performance filter: ${metricField} ${gaqlOperator} ${metricValue}`);
  }

  return whereConditions;
}

/**
 * Checks if a metric is a base metric that can be filtered in GAQL
 * @param {string} metricName The name of the metric
 * @return {boolean} True if it's a base metric
 * @private
 */
function _isBaseMetric(metricName) {
  return Object.keys(BASE_METRICS).some(key => key.split('.')[1] === metricName);
}

/**
 * Formats a metric for use in a GAQL query
 * @param {string} metricName The name of the metric
 * @param {number} thresholdValue The threshold value
 * @return {Object} Object with metricField and metricValue properties
 * @private
 */
function _formatMetricForQuery(metricName, thresholdValue) {
  // Special handling for cost which needs to be converted to micros
  if (metricName === 'cost') {
    return {
      metricField: 'metrics.cost_micros',
      metricValue: thresholdValue * 1000000 // Convert to micros
    };
  }

  return {
    metricField: `metrics.${metricName}`,
    metricValue: thresholdValue
  };
}

class SearchTermMatcher {
  constructor(debug = false) {
    this.debug = debug;
  }

  matchesCondition(term, condition) {
    if (!condition || !condition.text) {
      if (this.debug) console.warn('[SearchTermMatcher] Invalid condition:', condition);
      return false;
    }

    const termLower = term.toLowerCase();
    const textLower = condition.text.toLowerCase();
    let result = false;

    switch (condition.matchType) {
      case 'contains':
        result = termLower.includes(textLower);
        return result;

      case 'not-contains':
        result = !termLower.includes(textLower);
        return result;

      case 'regex-contains':
        try {
          const regex = new RegExp(condition.text, 'i');
          result = regex.test(term);
          return result;
        } catch (e) {
          if (this.debug) console.error('[SearchTermMatcher] Invalid regex:', e);
          return false;
        }

      case 'approx-contains':
        result = this.approxMatchContains(term, condition);
        return result;

      default:
        return false;
    }
  }

  approxMatchContains(term, condition) {
    if (!condition.threshold) {
      condition.threshold = 80; // Default threshold
    }

    const wholeTermScore = fuzzyMatchScore(condition.text, term);
    if (wholeTermScore >= condition.threshold) {
      return true;
    }
    const spacelessWholeTermScore = fuzzyMatchScore(condition.text.replace(/\s+/g, ''), term);
    if (spacelessWholeTermScore >= condition.threshold) {
      return true;
    }

    const termWords = term.toLowerCase().split(' ');
    const anyWordMatches = termWords.some(word => fuzzyMatchScore(condition.text, word) >= condition.threshold);
    if (anyWordMatches) {
      return true;
    }

    let startPosition = 0;
    let endPosition = condition.text.length;
    while (endPosition < term.length) {
      const score = fuzzyMatchScore(condition.text, term.substring(startPosition, endPosition));
      if (score >= condition.threshold) {
        return true;
      }
      startPosition++;
      endPosition++;
    }
    return false;
  }

  shouldExclude(term, rules) {
    // Parse rules if it's a string
    let parsedRules = rules;
    if (typeof rules === 'string') {
      try {
        parsedRules = JSON.parse(rules);
      } catch (e) {
        if (this.debug) console.error('[SearchTermMatcher] Error parsing rules string:', e);
        return false;
      }
    }

    // Check if we have valid sections
    if (!parsedRules?.sections?.length) {
      if (this.debug) console.log(`[SearchTermMatcher] No rules sections for term "${term}", not excluding`);
      return false;
    }

    let shouldExclude = false;

    // Loop through each section (sections are combined with AND logic)
    for (let i = 0; i < parsedRules.sections.length; i++) {
      const section = parsedRules.sections[i];

      // Skip invalid sections
      if (!section || !Array.isArray(section.conditions) || section.conditions.length === 0) {
        if (this.debug) console.log(`[SearchTermMatcher] Section ${i} is invalid or empty, skipping`);
        continue;
      }

      let sectionMatches = false;

      // Check if ANY condition in this section matches (OR logic within section)
      for (let j = 0; j < section.conditions.length; j++) {
        const condition = section.conditions[j];
        const conditionMatches = this.matchesCondition(term, condition);

        if (this.debug) console.log(`[SearchTermMatcher] Term "${term}" ${conditionMatches ? 'matches' : 'does not match'} condition [${i}][${j}]: ${condition.matchType} "${condition.text}"`);

        // If any condition matches, the section matches (OR logic)
        if (conditionMatches) {
          sectionMatches = true;
          break;
        }
      }

      if (this.debug) console.log(`[SearchTermMatcher] Section ${i} ${sectionMatches ? 'matches' : 'does not match'} term "${term}"`);

      // If a section doesn't match, the term should be excluded (failing AND logic)
      if (!sectionMatches) {
        shouldExclude = true;
        break;
      }
    }

    if (this.debug) console.log(`[SearchTermMatcher] Term "${term}" should be ${shouldExclude ? 'excluded' : 'included'}`);
    return shouldExclude;
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
  return result[0][0] * 100;
}

/**
 * Service class for sending email notifications about rule execution results.
 */
class EmailService {
  /**
   * Create a new EmailService instance.
   * @param {Object} options Configuration options for the email
   * @param {Object} options.rule The rule object
   * @param {Object} options.processingResult Results from processing the rule
   * @param {string} options.reportUrl URL to the generated report
   * @param {string} options.executionMode Mode of execution (Live, Preview, Test)
   * @param {string} options.ruleLevel Level of the rule (campaign or adgroup)
   * @param {string} options.timestamp Timestamp of execution
   * @param {string} options.accountName Google Ads account name
   * @param {string} options.accountId Google Ads account ID
   * @param {Array<string>} options.globalEmailRecipients Optional list of global email recipients
   * @param {string} options.emailLastSentTimestamp Last email sent timestamp
   * @param {Sheet} options.ruleSheet The sheet object for the rule
   */
  constructor(options) {
    this.rule = options.rule;
    this.processingResult = options.processingResult;
    this.reportUrl = options.reportUrl;
    this.executionMode = options.executionMode;
    this.ruleLevel = options.ruleLevel;
    this.timestamp = options.timestamp;
    this.accountName = options.accountName;
    this.accountId = options.accountId;
    this.globalEmailRecipients = options.globalEmailRecipients || [];
    this.emailLastSentTimestamp = options.emailLastSentTimestamp;
    this.ruleSheet = options.ruleSheet;
    // Combine rule-specific recipients with global recipients
    this.recipients = this._getAllRecipients(options.rule.email_recipients);
  }

  /**
   * Combine and deduplicate rule-specific and global email recipients.
   * @param {string} ruleRecipients Comma-separated email addresses from rule
   * @return {Array<string>} Combined and deduplicated array of email addresses
   * @private
   */
  _getAllRecipients(ruleRecipients) {
    // Parse rule-specific recipients
    const ruleEmailAddresses = this._parseRecipients(ruleRecipients);

    // If no rule-specific emails and no global emails, return empty array
    if (ruleEmailAddresses.length === 0 && this.globalEmailRecipients.length === 0) {
      return [];
    }

    // Combine rule-specific and global recipients
    const allRecipients = [...ruleEmailAddresses, ...this.globalEmailRecipients];

    // Deduplicate by converting to Set and back to Array
    return [...new Set(allRecipients)];
  }

  /**
   * Parse email recipients string into an array.
   * @param {string} recipientsString Comma-separated email addresses
   * @return {Array<string>} Array of email addresses
   * @private
   */
  _parseRecipients(recipientsString) {
    if (!recipientsString) return [];
    return recipientsString.split(',').map(email => email.trim()).filter(email => email);
  }

  /**
   * Send the email notification.
   */
  sendEmail() {
    if (this.recipients.length === 0) {
      console.log(`No valid email recipients found for rule "${this.rule.name}". Skipping email notification.`);
      return;
    }

    // Check if email has already been sent today
    const currentDateTime = new Date();

    // Check the last email timestamp for this specific rule
    const previousEmailTimestamp = this.emailLastSentTimestamp;

    if (previousEmailTimestamp) {
      // If an email was already sent today for this rule, skip sending
      if (this._hasSentEmailToday(previousEmailTimestamp)) {
        console.log(`Email already sent today (${previousEmailTimestamp}) for rule "${this.rule.name}". Skipping email notification.`);
        return;
      }
    }

    // Continue with sending email
    const subject = this._getSubject();
    const htmlBody = this._getHtmlBody();

    try {
      MailApp.sendEmail({
        to: this.recipients.join(','),
        subject: subject,
        htmlBody: htmlBody
      });
      const lastEmailTimestampRange = this.ruleSheet.getRange(RULE_LAST_EMAIL_VALUE_RANGE);
      // Update the timestamp after successful sending
      lastEmailTimestampRange.setValue(currentDateTime);

      console.log(`Email notification sent to ${this.recipients.join(', ')} for rule "${this.rule.name}".`);
    } catch (e) {
      console.log(`ERROR sending email notification for rule "${this.rule.name}": ${e.message}`);
    }
  }

  _hasSentEmailToday(previousEmailTimestamp) {
    const previousEmailDate = new Date(previousEmailTimestamp);
    if (!previousEmailDate) {
      return false;
    }
    const today = new Date();
    return today.toDateString() === previousEmailDate.toDateString();
  }

  /**
   * Generate the email subject line.
   * @return {string} The formatted subject line
   * @private
   */
  _getSubject() {
    const negativeCount = this.rule.numberOfResults;
    return `Auto-Negs: ${negativeCount} Results - ${this.rule.name} - ${this.accountName} (${this.executionMode})`;
  }

  /**
   * Generate the HTML body for the email.
   * @return {string} The formatted HTML email body
   * @private
   */
  _getHtmlBody() {
    const stats = this.processingResult.stats;
    const previewRows = this._getPreviewRows();
    const negativeCount = stats.collectedForNeg;
    const spreadsheetId = extractSpreadsheetId_(SPREADSHEET_URL);
    const editUrl = `https://autoneg.shabba.io/?id=${spreadsheetId}&ruleNumber=${this.rule.rule_number}`;

    return `
        <html>
        <head>
            <style>
                body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto; }
                .header { background-color: #4285f4; color: white; padding: 20px; border-radius: 5px 5px 0 0; }
                .content { padding: 20px; border: 1px solid #ddd; border-top: none; border-radius: 0 0 5px 5px; }
                .section { margin-bottom: 30px; }
                h1 { margin: 0; font-size: 24px; }
                h2 { margin-top: 0; color: #4285f4; font-size: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px; }
                table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
                th, td { text-align: left; padding: 12px; border-bottom: 1px solid #ddd; }
                th { background-color: #f2f2f2; }
                
                /* Results Summary Stats styling */
                .results-summary { 
                    margin: 25px 0;
                    text-align: center;
                }
                .stats-container {
                    display: flex;
                    justify-content: space-between;
                    flex-wrap: wrap;
                    gap: 15px;
                    padding: 20px;
                    background-color: #f9f9f9;
                    border-radius: 8px;
                }
                .stat-box {
                    flex: 1;
                    min-width: 150px;
                    background: white;
                    padding: 20px 15px;
                    border-radius: 8px;
                    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
                }
                .stat-value {
                    font-size: 32px;
                    font-weight: bold;
                    color: #4285f4;
                    margin-bottom: 8px;
                }
                .stat-label {
                    font-size: 14px;
                    color: #666;
                }
                
                .button {
                    display: inline-block;
                    background-color: #4285f4;
                    color: white !important;
                    padding: 10px 20px;
                    text-decoration: none;
                    border-radius: 5px;
                    font-weight: bold;
                    margin: 5px;
                }
                .footer {
                    margin-top: 30px;
                    font-size: 12px;
                    color: #666;
                    border-top: 1px solid #eee;
                    padding-top: 20px;
                    text-align: center;
                }
                .mode-live { color: #34a853; }
                .mode-preview { color: #fbbc05; }
                .mode-test { color: #ea4335; }
                
                .action-buttons {
                    text-align: center;
                    margin: 20px 0;
                }
                
                .edit-rule-link {
                    display: inline-block;
                    margin-top: 10px;
                    color: #4285f4;
                    text-decoration: underline;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Auto-Negs Rule Results</h1>
            </div>
            <div class="content">
                <div class="section">
                    <h2>Rule Information</h2>
                    <table>
                        <tr><td><strong>Rule Name:</strong></td><td>${this.rule.name}</td></tr>
                        <tr><td><strong>Rule #:</strong></td><td>${this.rule.rule_number}</td></tr>
                        <tr><td><strong>Account:</strong></td><td>${this.accountName} (${this.accountId})</td></tr>
                        <tr><td><strong>Mode:</strong></td><td><span class="mode-${this.executionMode.toLowerCase().replace(' ', '-')}">${this.executionMode}</span></td></tr>
                        <tr><td><strong>Level:</strong></td><td>${this._capitalizeFirstLetter(this.ruleLevel)}</td></tr>
                        <tr><td><strong>Run Date:</strong></td><td>${this.timestamp}</td></tr>
                    </table>
                </div>
                
                <div class="section">
                    <h2>Results Summary</h2>
                    <div class="results-summary">
                        <div class="stats-container">
                            <div class="stat-box">
                                <div class="stat-value">${stats.checked}</div>
                                <div class="stat-label">Search Terms Checked</div>
                            </div>
                            <div class="stat-box">
                                <div class="stat-value">${stats.matched}</div>
                                <div class="stat-label">Matching Filters</div>
                            </div>
                            <div class="stat-box">
                                <div class="stat-value">${negativeCount}</div>
                                <div class="stat-label">Negatives ${this._getNegativeActionText()}</div>
                            </div>
                        </div>
                    </div>
                </div>
                
                ${previewRows.length > 0 ? this._getPreviewTableHtml(previewRows) : this._getNoResultsHtml()}
                
                <div class="action-buttons">
                    ${this.reportUrl ? `<a href="${this.reportUrl}" class="button">View Full Report</a>` : ''}
                    <a href="${editUrl}" class="button">Edit Rule</a>
                </div>
                
                <div class="footer">
                    <p>This email was automatically sent by the Auto Negative Keywords Scripts.</p>
                    <p>Contact Charles Bannister for feedback, support or customisations. (<a href="https://www.linkedin.com/in/charles-bannister-92a1a228/">LinkedIn</a>) (<a href="mailto:charles@shabba.io">Email</a>)</p>
                    <p>Report generated on ${this.timestamp}.</p>
                </div>
            </div>
        </body>
        </html>
        `;
  }

  /**
   * Get the preview rows for the email (limited to 10 rows).
   * @return {Array<Array>} Array of row data for the preview table
   * @private
   */
  _getPreviewRows() {
    const maxPreviewRows = 10;
    return this.processingResult.outputDataRows.slice(0, maxPreviewRows);
  }

  /**
   * Generate HTML for the preview table.
   * @param {Array<Array>} previewRows Array of row data
   * @return {string} HTML for the preview table section
   * @private
   */
  _getPreviewTableHtml(previewRows) {
    if (previewRows.length === 0) return '';

    // Get headers based on rule level
    let headers = [
      "Search Term", "Added/Removed",
      "Impressions", "Clicks", "Cost", "Conversions", "Conv. Value",
      "CTR", "Conv. Rate", "Cost/Conv.", "CPC", "ROAS", "Negative Added", "AI Response"
    ];

    if (this.ruleLevel === 'campaign' || this.ruleLevel === 'adgroup') {
      headers.splice(1, 0, "Campaign Name");
    }
    if (this.ruleLevel === 'adgroup') {
      headers.splice(2, 0, "Ad Group Name");
    }

    // Build table HTML
    let tableHtml = `
        <div class="section">
            <h2>Preview of Results</h2>
            <table>
                <tr>
                    ${headers.map(header => `<th>${header}</th>`).join('')}
                </tr>
        `;

    // Add rows
    previewRows.forEach(row => {
      tableHtml += `
                <tr>
                    ${row.map(cell => `<td>${cell}</td>`).join('')}
                </tr>
            `;
    });

    tableHtml += `
            </table>
            ${this.processingResult.outputDataRows.length > previewRows.length ?
        `<p><em>Showing ${previewRows.length} of ${this.processingResult.outputDataRows.length} results. View the full report for all results.</em></p>` : ''}
        </div>
        `;

    return tableHtml;
  }

  /**
   * Generate HTML for the no results message.
   * @return {string} HTML for the no results section
   * @private
   */
  _getNoResultsHtml() {
    return `
        <div class="section" style="text-align: center; padding: 30px; background-color: #f9f9f9; border-radius: 5px;">
            <p style="font-size: 18px;">No matching search terms found for this rule.</p>
        </div>
        `;
  }

  /**
   * Get the appropriate text for what happened to negatives based on execution mode.
   * @return {string} Text describing what happened to negatives
   * @private
   */
  _getNegativeActionText() {
    if (this.executionMode === "Live Mode") {
      return "Added";
    } else if (this.executionMode === "Preview Mode") {
      return "Previewed";
    } else { // Test Mode
      return "Found";
    }
  }

  /**
   * Capitalize the first letter of a string.
   * @param {string} string The input string
   * @return {string} The string with first letter capitalized
   * @private
   */
  _capitalizeFirstLetter(string) {
    return string.charAt(0).toUpperCase() + string.slice(1);
  }
}

/**
 * Gets global email recipients from the Settings sheet.
 * @param {Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @return {Array<string>} Array of email addresses or empty array if not found.
 * @private
 */
function getGlobalEmailRecipients_(spreadsheet) {
  try {
    const settingsSheet = spreadsheet.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      console.log(`Settings sheet "${SETTINGS_SHEET_NAME}" not found. No global email recipients will be used.`);
      return [];
    }

    const emailsCell = settingsSheet.getRange(SETTINGS_EMAIL_RANGE);
    const emailsString = emailsCell.getValue().toString().trim();

    if (!emailsString) {
      console.log(`No email addresses found in Settings sheet cell ${SETTINGS_EMAIL_RANGE}.`);
      return [];
    }

    // Split by comma, trim whitespace, and filter out empty entries
    return emailsString.split(',').map(email => email.trim()).filter(email => email);
  } catch (e) {
    console.log(`ERROR retrieving global email recipients: ${e.message}`);
    return [];
  }
}

/**
 * Checks if entity conditions include Performance Max campaigns.
 * @param {Array<Object>} entityConditions The entity conditions.
 * @return {boolean} True if PMAX campaigns are included, false otherwise.
 * @private
 */
function hasPmaxCampaignCondition_(entityConditions) {
  if (!Array.isArray(entityConditions)) return false;

  // Look for conditions that might include PMAX campaigns
  for (const condition of entityConditions) {
    if (condition.field === 'campaign.advertising_channel_type' &&
      (condition.value === 'PERFORMANCE_MAX' ||
        (condition.operator === 'contains' && condition.value.includes('PERFORMANCE_MAX')))) {
      return true;
    }

    // Check for campaign name conditions that might target PMAX
    if (condition.field === 'campaign.name' &&
      condition.operator === 'contains' &&
      condition.value.toLowerCase().includes('pmax')) {
      return true;
    }
  }

  return false;
}

/**
 * Class to handle building reports for Performance Max campaigns.
 * Gets all campaign IDs matching filters and fetches search terms for each.
 */
class PMaxReportBuilder {
  /**
   * Create a new PMaxReportBuilder instance.
   * @param {number} lookbackDays Number of days to look back.
   * @param {Array<Object>} entityConditions Entity conditions from the rule.
   */
  constructor(lookbackDays, entityConditions) {
    this.lookbackDays = lookbackDays;
    this.entityConditions = entityConditions;
  }

  /**
   * Find PMAX campaigns that match the entity conditions.
   * @return {Array<Object>} Array of matching PMAX campaign objects with id and name.
   */
  getMatchingPmaxCampaigns() {
    try {
      // Build a query to get all PMAX campaign IDs
      const dateCondition = _getDateRangeCondition_(this.lookbackDays);
      const statusConditions = ['campaign.status = ENABLED'];

      // Extract campaign-specific conditions
      const campaignConditions = [];

      if (Array.isArray(this.entityConditions)) {
        for (const condition of this.entityConditions) {
          if (condition.field && condition.field.startsWith('campaign.') &&
            !condition.field.startsWith('campaign.advertising_channel_type')) {
            const conditionString = _getEntityConditionString_(condition);
            if (conditionString) {
              campaignConditions.push(conditionString);
            }
          }
        }
      }

      // Always add the PMAX condition
      campaignConditions.push("campaign.advertising_channel_type = 'PERFORMANCE_MAX'");

      // Build the complete query
      const whereConditions = [dateCondition, ...statusConditions, ...campaignConditions];
      const query = `
                SELECT 
                    campaign.id,
                    campaign.name
                FROM campaign
                WHERE ${whereConditions.join(' AND ')}
            `;

      console.log(`Executing PMAX campaigns query: ${query}`);
      const report = AdsApp.report(query);
      const rows = report.rows();

      const campaigns = [];
      while (rows.hasNext()) {
        const row = rows.next();
        campaigns.push({
          id: row['campaign.id'],
          name: row['campaign.name']
        });
      }

      console.log(`Found ${campaigns.length} matching PMAX campaigns`);
      return campaigns;

    } catch (e) {
      console.log(`Error finding matching PMAX campaigns: ${e.message}`);
      return [];
    }
  }

  /**
   * Get PMAX search terms for the provided campaigns.
   * @param {Array<Object>} campaigns Array of campaign objects with id and name.
   * @return {Array<Object>} Array of PMAX search term objects.
   */
  getPmaxSearchTerms(campaigns) {
    if (!campaigns || !campaigns.length) {
      return [];
    }

    const allSearchTerms = [];

    // Process each campaign one at a time
    for (const campaign of campaigns) {
      try {
        const campaignId = campaign.id;
        const campaignName = campaign.name;
        const dateCondition = _getDateRangeCondition_(this.lookbackDays);

        // Build query to get PMAX search terms for this campaign
        // Note: campaign_search_term_insight only supports specific metrics
        const query = `
                    SELECT
                        campaign_search_term_insight.category_label,
                        metrics.clicks,
                        metrics.impressions,
                        metrics.conversions,
                        metrics.conversions_value
                    FROM campaign_search_term_insight
                    WHERE ${dateCondition}
                    AND campaign_search_term_insight.campaign_id = '${campaignId}'
                `;

        console.log(`Executing PMAX search terms query for campaign ${campaignId} (${campaignName})`);
        const report = AdsApp.report(query);
        const rows = report.rows();

        let termCount = 0;
        while (rows.hasNext()) {
          const row = rows.next();
          termCount++;

          if (row['campaign_search_term_insight.category_label'].trim() === '') {
            continue;
          }

          // Log if conversions_value is missing
          if (typeof row['metrics.conversions_value'] === 'undefined') {
            console.log(`Warning: PMAX search term "${row['campaign_search_term_insight.category_label']}" is missing conversions_value. Using default of 0.`);
          }

          // Format data in a way similar to regular search term report
          // For PMAX, we have to estimate cost since cost_micros isn't available
          // We'll use 0 for cost since it's not available directly
          allSearchTerms.push({
            searchTerm: row['campaign_search_term_insight.category_label'],
            campaignId: campaignId,
            campaignName: campaignName,
            impressions: parseInt(row['metrics.impressions'] || 0),
            clicks: parseInt(row['metrics.clicks'] || 0),
            cost: 0, // Cost not available for PMAX search terms
            conversions: parseFloat(row['metrics.conversions'] || 0),
            conversions_value: parseFloat(row['metrics.conversions_value'] || 0),
            isPmax: true
          });
        }

        console.log(`Retrieved ${termCount} search terms for PMAX campaign ${campaignId} (${campaignName})`);

      } catch (e) {
        console.log(`Error getting search terms for PMAX campaign ${campaign.id}: ${e.message}`);
        // Continue with the next campaign
      }
    }

    return allSearchTerms;
  }
}


/**
 * Normalise a LIKE or NOT LIKE string so that it works with the LIKE operator
 * From the docs:
 * To match a literal [, ], %, or _ using the LIKE operator, surround the character in square brackets.
 * For example, the following condition matches all campaign.name values that start with [Earth_to_Mars]:
 * campaign.name LIKE '[[]Earth[_]to[_]Mars[]]'
 */
function normaliseLikeString(string) {
  // Escape LIKE pattern special characters: _, [
  // Note: ] is not typically a special character on its own in LIKE,
  // but escaping it ensures correctness if it's part of a character set []
  // and prevents issues if future SQL versions change behavior.
  // We replace the character 'char' with '[char]'.
  const normalisedString = string.replace(/[_\[\]]/g, (char) => `[${char}]`);
  return normalisedString;
}

/**
 * Sorts rules by their last runtime from the spreadsheet.
 * Rules with no runtime (never run before) will be placed first.
 * Otherwise, rules are sorted by timestamp with the most recently run rule last.
 * 
 * @param {Array<Object>} rules The rules to sort.
 * @param {Spreadsheet} spreadsheet The Google Spreadsheet object.
 * @return {Array<Object>} The sorted rules array.
 * @private
 */
function sortRulesByLastRuntime_(rules, spreadsheet) {
  if (!rules || rules.length <= 1) {
    return rules; // Nothing to sort
  }

  // Create a map to store the last runtime for each rule
  const runtimeMap = {};

  // Loop through each rule to get its last runtime
  rules.forEach(rule => {
    try {
      // Get the sheet for the rule (sheet name may be prefixed for numeric IDs)
      const formattedSheetName = isNumeric(rule.rule_number) ? `${rule.rule_number}` : String(rule.rule_number);
      const sheet = spreadsheet.getSheetByName(formattedSheetName);

      if (!sheet) {
        // No sheet exists yet, this rule has never been run
        runtimeMap[rule.id] = null;
        return;
      }

      // Get the timestamp from cell A3
      const timestampCell = sheet.getRange("A3");
      const timestampValue = timestampCell.getValue();

      if (!timestampValue) {
        // No timestamp in the cell
        runtimeMap[rule.id] = null;
        return;
      }

      // Convert to Date object if it's a string
      const timestamp = (typeof timestampValue === 'string')
        ? new Date(timestampValue)
        : timestampValue; // Assume it's already a date object

      // Store the timestamp (we'll use getTime() for sorting)
      runtimeMap[rule.id] = timestamp;

    } catch (e) {
      console.log(`Error getting runtime for rule ${rule.id}: ${e.message}`);
      runtimeMap[rule.id] = null; // Mark as no runtime in case of error
    }
  });

  // Sort the rules based on their last runtime
  return rules.sort((a, b) => {
    const timeA = runtimeMap[a.id] ? runtimeMap[a.id].getTime() : -Infinity;
    const timeB = runtimeMap[b.id] ? runtimeMap[b.id].getTime() : -Infinity;

    // Rules with no runtime (null) go first
    if (timeA === -Infinity && timeB !== -Infinity) return -1;
    if (timeA !== -Infinity && timeB === -Infinity) return 1;

    // Otherwise, sort by timestamp (older to newer)
    return timeA - timeB;
  });
}

/**
 * Fetches Performance Max search terms data for a rule
 * @param {string} ruleName Name of the rule
 * @param {number} lookbackDays Number of days to look back
 * @param {Array} entityConditions Conditions from the rule
 * @returns {Array} Array of PMAX search term rows or empty array if none found
 * @private
 */
function _fetchPmaxData_(ruleName, lookbackDays, entityConditions) {
  let pmaxRows = [];

  try {
    // Create PMAX report builder
    const pmaxReportBuilder = new PMaxReportBuilder(lookbackDays, entityConditions);

    // First check if any PMAX campaigns match the conditions
    const matchingCampaigns = pmaxReportBuilder.getMatchingPmaxCampaigns();

    if (matchingCampaigns && matchingCampaigns.length > 0) {
      console.log(`Rule "${ruleName}" matched ${matchingCampaigns.length} Performance Max campaigns. Fetching search terms...`);

      // Get search terms for the matching campaigns
      pmaxRows = pmaxReportBuilder.getPmaxSearchTerms(matchingCampaigns);

      if (pmaxRows && pmaxRows.length > 0) {
        console.log(`Found ${pmaxRows.length} PMAX search terms for rule "${ruleName}".`);
      } else {
        console.log(`No PMAX search terms found for matching campaigns in rule "${ruleName}".`);
      }
    }
  } catch (e) {
    console.log(`ERROR executing PMAX report for rule "${ruleName}": ${e.message}. Continuing with regular report.`);
  }

  return pmaxRows;
}

/**
 * 
 * @param {String} searchTerm 
 * @param {Object} rule 
 * @returns {String} The AI response
 */
function getAiResponseString_(searchTerm, rule, campaignName, adGroupName, campaignId, adGroupId) {
  if (rule.ai_prompt.trim() === '') {
    return "N/A (no prompt defined)";
  }
  const prompt = interpolatePrompt(searchTerm, rule, campaignName, adGroupName, campaignId, adGroupId);
  const { llmName, modelName, apiKey, logPrompts } = getLlmConfig();
  const { responseGetter } = getLlmStrategy(llmName);
  if (logPrompts) {
    console.log(`LLM: ${llmName} - Prompt: \n${prompt}\n`);
  }
  // log the first prompt regardless of the logPrompts setting
  if (!FIRST_PROMPT_LOGGED) {
    console.log(`\nLLM: ${llmName} - Prompt: \n${prompt}\n\n`);
    FIRST_PROMPT_LOGGED = true;
  }
  const response = aiResponse(responseGetter, modelName, prompt, apiKey);
  if (logPrompts) {
    console.log(`LLM: ${llmName} - Response: \n${response}\n`);
  }
  return response;
}

//get llm config from the settings sheet
function getLlmConfig() {
  const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  const settingsSheet = spreadsheet.getSheetByName('Settings');
  const llmName = settingsSheet.getRange('B9').getValue().toUpperCase();
  const defaultModelName = DEFAULT_MODEL_NAMES[llmName];
  const modelName = settingsSheet.getRange('B10').getValue() || defaultModelName;
  const apiKey = settingsSheet.getRange('B11').getValue();
  const logPrompts = settingsSheet.getRange(SETTINGS_LOG_PROMPTS_RANGE).getValue() || false;
  return { llmName, modelName, apiKey, logPrompts };
}

function aiResponse(fn, modelName, promptBody, apiKey) {
  try {
    Utilities.sleep(LLM_RESPONSE_DELAY_TIME_MILLISECONDS);
    const response = fn(modelName, promptBody, apiKey);
    // Remove special characters and make lowercase
    const strippedResponse = response.replace(/[^a-zA-Z0-9\s]/g, '').toLowerCase();
    return strippedResponse;
  } catch (e) {
    console.error(`There was a problem with the AI response. Returing "N/A". Error: ${e}`);
    console.error(`Did you set a valid API key in the Settings sheet?`);
    return "N/A";
  }
}


function getAwanResponse(modelName, promptBody, apiKey) {
  const url = "https://api.awanllm.com/v1/completions"
  const payload = {
    "model": modelName,
    "prompt": `<|begin_of_text|><|start_header_id|>system<|end_header_id|>\n\n${promptBody}<|eot_id|><|start_header_id|>user<|end_header_id|>`,
    "repetition_penalty": 1.1,
    "temperature": 0.7,
    "top_p": 0.9,
    "top_k": 40,
    "max_tokens": 1024,
    "stream": true
  }
  const headers = {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${apiKey}`
  };

  const text = apiCall(parseAWANResponse, url, payload, headers);

  return text;

}

function getChatGptResponse(modelName, promptBody, apiKey) {
  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    "model": modelName,
    "messages": [{
      "role": "user",
      "content": promptBody
    }]
  }
  const headers = {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${apiKey}`
  }
  const response = apiCall(parseChatGptResponse, url, payload, headers);
  return response;
}

function parseChatGptResponse(response) {
  const parsedResult = JSON.parse(response);
  return parsedResult.choices[0].message.content;
}

function parseAWANResponse(response) {
  const data = response.replace('\n\ndata: [DONE]\n\n', '')
    .split('\n')
    .filter(x => x !== '')
    .map(x => x.replace('data: ', ''))
    .map(x => {
      let formattedJson = x.replace(/'/g, '"');
      formattedJson = formattedJson.replace('Array(1)', '[]').replace('…', 'null');
      try {
        return JSON.parse(formattedJson);
      } catch (e) {
        return JSON.parse(x);
      }
    });
  const text = data.map(x => x.choices[0].text).join('\n').replace(/\n/g, '');
  return text;
}

function getClaudeResponse(modelName, promptBody, apiKey) {
  const url = "https://api.anthropic.com/v1/messages";
  const payload = {
    "model": `${modelName}`,
    "max_tokens": 1024,
    "messages": [{
      "role": 'user',
      "content": promptBody
    }]
  }
  const headers = {
    'Content-Type': 'application/json',
    'x-api-key': `${apiKey}`,
    'anthropic-version': '2023-06-01'
  };

  const parseResult = result => {
    const parsedResult = JSON.parse(result);
    return parsedResult.content[0].text;
  };

  const text = apiCall(parseResult, url, payload, headers);

  return text;

}


function getGeminiResponse(modelName, promptBody, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  const payload = {
    "contents": [{
      "parts": [{ "text": promptBody }]
    }]
  };
  const headers = {
    'Content-Type': 'application/json',
  };
  const parseGeminiResponse = (response) => {
    const parsedResult = JSON.parse(response);
    // Adjust parsing for the potentially nested response structure
    // Check if candidates exist and have the expected structure
    if (parsedResult.candidates && parsedResult.candidates.length > 0 &&
      parsedResult.candidates[0].content && parsedResult.candidates[0].content.parts &&
      parsedResult.candidates[0].content.parts.length > 0 && parsedResult.candidates[0].content.parts[0].text) {
      return parsedResult.candidates[0].content.parts[0].text;
    } else {
      // Log the full response if the structure is unexpected
      console.error("Unexpected Gemini response structure:", JSON.stringify(parsedResult, null, 2));
      // Handle potential errors, like blocked prompts
      if (parsedResult.promptFeedback && parsedResult.promptFeedback.blockReason) {
        return `Blocked: ${parsedResult.promptFeedback.blockReason}`;
      }
      return "Error: Could not parse Gemini response.";
    }
  }
  const response = apiCall(parseGeminiResponse, url, payload, headers);
  return response;
}


function apiCall(responseParser, url, payload, headers) {
  const options = {
    'method': 'POST',
    'headers': headers,
    'payload': JSON.stringify(payload)
  };

  return apiCallWithRetry(options, url, responseParser);
}

function apiCallWithRetry(options, url, responseParser, attempt = 1, maxRetries = 5) {
  try {
    const response = UrlFetchApp.fetch(url, options);
    const contextText = response.getContentText();
    const text = responseParser(contextText);
    return text;
  } catch (error) {
    const shouldRetry = isRateLimitError(error, attempt, maxRetries);

    if (!shouldRetry) {
      throw error;
    }

    const delay = calculateExponentialBackoff(attempt);
    console.log(`Rate limit error detected. Retrying in ${delay}ms (attempt ${attempt}/${maxRetries})`);
    Utilities.sleep(delay);

    return apiCallWithRetry(options, url, responseParser, attempt + 1, maxRetries);
  }
}

function isRateLimitError(error, attempt, maxRetries) {
  if (attempt >= maxRetries) {
    return false;
  }

  const errorString = error.toString().toLowerCase();
  const errorMessage = error.message ? error.message.toLowerCase() : '';

  // Check for 429 status code
  if (errorString.includes('429') || errorMessage.includes('429')) {
    return true;
  }

  // Check for rate limit related error messages
  const rateLimitIndicators = [
    'rate limit',
    'rate_limit_error',
    'too many requests',
    'quota exceeded',
    'request limit exceeded'
  ];

  return rateLimitIndicators.some(indicator =>
    errorString.includes(indicator) || errorMessage.includes(indicator)
  );
}

function calculateExponentialBackoff(attempt) {
  // Base delay of 1 second, doubled for each attempt with jitter
  const baseDelay = 1000;
  const exponentialDelay = baseDelay * Math.pow(2, attempt - 1);
  // Add jitter (random factor between 0.5 and 1.5) to avoid thundering herd
  const jitter = 0.5 + Math.random();
  return Math.floor(exponentialDelay * jitter);
}


function interpolatePrompt(searchTerm, rule, campaignName, adGroupName, campaignId, adGroupId) {

  const prompt = rule.ai_prompt;

  let interpolatedPrompt = prompt.replace('{{search_term}}', searchTerm)
    .replace('{{campaign_name}}', campaignName)
    .replace('{{ad_group_name}}', adGroupName);
  if (rule.ai_prompt.includes('{{keywords}}')) {
    const keywords = getKeywords(campaignId, adGroupId);
    interpolatedPrompt = interpolatedPrompt.replace('{{keywords}}', keywords);
  }
  if (rule.ai_prompt.includes('{{exact_keywords}}')) {
    const exactKeywords = getKeywords(campaignId, adGroupId, 'EXACT');
    interpolatedPrompt = interpolatedPrompt.replace('{{exact_keywords}}', exactKeywords);
  }
  if (rule.ai_prompt.includes('{{phrase_keywords}}')) {
    const phraseKeywords = getKeywords(campaignId, adGroupId, 'PHRASE');
    interpolatedPrompt = interpolatedPrompt.replace('{{phrase_keywords}}', phraseKeywords);
  }

  return interpolatedPrompt;
}

/**
 * Get keywords for a campaign or ad group
 * @param {string} campaignId The id of the campaign
 * @param {string} adGroupId The id of the ad group
 * @param {string} matchType The match type of the keyword
 * @returns {string} The keywords
 */
function getKeywords(campaignId = null, adGroupId = null, matchType = null) {
  let keywordGaqlQuery = `
		SELECT
			ad_group_criterion.keyword.text
		FROM keyword_view
		WHERE campaign.status = 'ENABLED'
		AND ad_group.status = 'ENABLED'
		AND ad_group_criterion.status = 'ENABLED'
		AND ad_group_criterion.negative = FALSE
		`;
  if (campaignId) {
    keywordGaqlQuery += ` AND campaign.id = '${campaignId}'`;
  }
  if (adGroupId) {
    keywordGaqlQuery += ` AND ad_group.id = '${adGroupId}'`;
  }
  if (matchType) {
    keywordGaqlQuery += ` AND keyword.match_type = '${matchType}'`;
  }
  // console.log(`Keyword GAQL Query: ${keywordGaqlQuery}`);
  const report = AdsApp.report(keywordGaqlQuery);
  const rows = report.rows();
  const keywords = [];
  while (rows.hasNext()) {
    const row = rows.next();
    if (keywords.includes(row['ad_group_criterion.keyword.text'])) {
      continue;
    }
    keywords.push(row['ad_group_criterion.keyword.text']);
  }
  return keywords.join(',');
}

function getLlmStrategy(llmName) {
  const llmStrategies = {
    [LLM_NAME_ENUM.CLAUDE]: {
      responseGetter: getClaudeResponse,
    },
    [LLM_NAME_ENUM.AWAN]: {
      responseGetter: getAwanResponse,
    },
    [LLM_NAME_ENUM.CHAT_GPT]: {
      responseGetter: getChatGptResponse,
    },
    [LLM_NAME_ENUM.GEMINI]: {
      responseGetter: getGeminiResponse,
    },
  }
  if (!llmStrategies[llmName]) {
    throw new Error(`Invalid LLM name: ${llmName}`);
  }
  return llmStrategies[llmName];
}

/**
 * Checks if a value is numeric.
 * @param {*} value The value to check.
 * @return {boolean} True if the value is numeric.
 * @private
 */
function isNumeric(value) {
  return !isNaN(parseInt(value)) && String(value).trim() === String(parseInt(value));
}