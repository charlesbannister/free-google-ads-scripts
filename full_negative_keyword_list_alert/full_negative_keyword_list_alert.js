/*
Name: Negative Keyword List Full Checker
Summary: Checks for full shared negative keyword lists and alerts once via email
@Author: Charles Bannister
More scripts at https://shabba.io
Version: 1.0.0
*/

// CONFIG
const DEBUG_MODE = false;
const NEGATIVE_KEYWORD_LIMIT = 5000; // Max number of keywords per list
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';
// Used to track which lists have already triggered alerts. Create a new sheet (sheets.new) and paste in the URL
const ALERT_EMAIL = 'YOUR_EMAIL_HERE';
// Email to notify when lists hit the threshold

/**
 * Main entry point
 */
function main() {
    console.log('Script started');

    const negativeKeywordLists = getNegativeKeywordListData();
    logNegativeKeywordListSummary(negativeKeywordLists);

    const alertedListIds = getAlreadyAlertedListIds();
    const fullListsToAlert = getNewFullListsToAlert(negativeKeywordLists, alertedListIds);

    logListsToAlert(fullListsToAlert);

    if (fullListsToAlert.length > 0) {
        sendAlertEmail(fullListsToAlert);
        storeAlertedLists(fullListsToAlert);
    }

    console.log('Script finished');
}

/**
 * Retrieves shared negative keyword list data and keyword counts
 * @returns {Array<Object>} Array of flat objects with list name, ID, and keyword count
 */
function getNegativeKeywordListData() {
    const query = getNegativeKeywordListGaqlQuery();
    const report = AdsApp.report(query);
    const rows = report.rows();

    const keywordListMap = {};

    while (rows.hasNext()) {
        const row = rows.next();
        const listId = row['shared_criterion.shared_set'];
        const listName = row['shared_set.name'];

        if (!keywordListMap[listId]) {
            keywordListMap[listId] = { listId, listName, keywordCount: 0 };
        }

        keywordListMap[listId].keywordCount += 1;
    }

    return Object.values(keywordListMap);
}

/**
 * Builds the GAQL query to get shared negative keyword list members
 * @returns {string} GAQL query
 */
function getNegativeKeywordListGaqlQuery() {
    const query = `
    SELECT
      shared_criterion.shared_set,
      shared_set.name
    FROM shared_criterion
    WHERE shared_set.type = 'NEGATIVE_KEYWORDS'
  `;

    console.log('GAQL Query used:');
    console.log(query);
    return query;
}

/**
 * Logs a summary of the keyword lists (first 3)
 * @param {Array<Object>} listData - Array of keyword list objects
 */
function logNegativeKeywordListSummary(listData) {
    if (!listData.length) {
        console.log('No negative keyword lists found.');
        return;
    }

    const sample = listData.slice(0, 3);

    console.log(`\nSummary of first 3 negative keyword lists:`);
    sample.forEach((list, index) => {
        console.log(`\nList ${index + 1}:
  Name: ${list.listName}
  ID: ${list.listId}
  Keyword Count: ${list.keywordCount}`);
    });
    console.log('');
}

/**
 * Gets the sheet and returns all previously alerted list IDs
 * @returns {Array<string>} List of alerted list IDs
 */
function getAlreadyAlertedListIds() {
    if (SPREADSHEET_URL === 'PASTE_YOUR_SPREADSHEET_URL_HERE') {
        throw new Error('Please set the SPREADSHEET_URL to your tracking sheet URL.');
    }

    const sheet = getOrCreateSheet('Alerted Lists');
    const data = sheet.getDataRange().getValues();
    const listIds = data.slice(1).map(row => row[0]);

    return listIds;
}

/**
 * Filters for full lists that haven't been alerted
 * @param {Array<Object>} allLists - All keyword list data
 * @param {Array<string>} alertedIds - Already alerted list IDs
 * @returns {Array<Object>} Filtered list to alert
 */
function getNewFullListsToAlert(allLists, alertedIds) {
    return allLists.filter(list => {
        return list.keywordCount >= NEGATIVE_KEYWORD_LIMIT && !alertedIds.includes(list.listId);
    });
}

/**
 * Logs lists that will be alerted
 * @param {Array<Object>} listsToAlert - Keyword lists to alert on
 */
function logListsToAlert(listsToAlert) {
    if (!listsToAlert.length) {
        console.log('No new full lists to alert.');
        return;
    }

    console.log(`\nLists that are full and need alerting:`);
    listsToAlert.forEach(list => {
        console.log(`\nName: ${list.listName}
ID: ${list.listId}
Keyword Count: ${list.keywordCount}`);
    });
    console.log('');
}

/**
 * Sends an alert email with full list details
 * @param {Array<Object>} listsToAlert - Lists that are full
 */
function sendAlertEmail(listsToAlert) {
    const subject = '⚠️ Full Negative Keyword List(s) Detected';
    let body = '<p>The following negative keyword lists have reached the limit of ' + NEGATIVE_KEYWORD_LIMIT + ' keywords:</p><ul>';

    listsToAlert.forEach(list => {
        body += `<li><strong>${list.listName}</strong> (ID: ${list.listId}) – ${list.keywordCount} keywords</li>`;
    });

    body += '</ul><p>This alert will only be sent once per list.</p>';

    MailApp.sendEmail({
        to: ALERT_EMAIL,
        subject: subject,
        htmlBody: body
    });

    console.log(`Alert email sent to ${ALERT_EMAIL}`);
}

/**
 * Stores the alerted list IDs and names in the sheet
 * @param {Array<Object>} alertedLists - Lists that were just alerted
 */
function storeAlertedLists(alertedLists) {
    const sheet = getOrCreateSheet('Alerted Lists');
    const now = new Date();

    alertedLists.forEach(list => {
        sheet.appendRow([list.listId, list.listName, now]);
    });

    console.log(`${alertedLists.length} list(s) recorded in sheet.`);
}

/**
 * Gets or creates a sheet by name
 * @param {string} sheetName - Sheet tab name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Sheet object
 */
function getOrCreateSheet(sheetName) {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    let sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        sheet.appendRow(['List ID', 'List Name', 'Alert Date']);
        sheet.getRange('1:1').setFontWeight('bold');
        sheet.setFrozenRows(1);
    }

    return sheet;
}
