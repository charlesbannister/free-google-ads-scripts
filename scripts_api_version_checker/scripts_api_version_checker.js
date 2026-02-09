/**
 * Scripts API Version Checker
 * Google Ads Scripts usually lag behind Google Ads API versions
 * Use this to get an alert when the version updates
 * Version: 1.0.0
 * @author: Charles Bannister of shabba.io
 */

// Email address to receive version update notifications
const EMAIL = 'YOUR_EMAIL_HERE';

// Spreadsheet URL where the version will be stored in cell A1 of the first sheet
// Create a new google sheet by typing "sheets.new" into the address bar
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';

// Email recipients list
const RECIPIENTS = [EMAIL];

// Email subject line
const SUBJECT = 'Google Ads API for Scripts has been upgraded';

// URL to Google Ads API release notes
const RELEASE_NOTES_URL = 'https://developers.google.com/google-ads/api/docs/release-notes';

// Cell reference where version is stored (A1)
const VERSION_CELL = 'A1';

function main() {
  console.log('Script started');

  const errorMessages = checkVersion();
  console.log(`Error messages received: ${JSON.stringify(errorMessages)}`);

  if (!errorMessages || errorMessages.length === 0) {
    console.warn('No error messages returned. Unable to determine API version.');
    return;
  }

  const currentVersion = extractVersionFromError(errorMessages[0]);
  console.log(`Extracted current version: ${currentVersion}`);

  if (!currentVersion) {
    console.warn('Could not extract version from error message.');
    return;
  }

  const storedVersion = getStoredVersion();
  console.log(`Stored version from sheet: ${storedVersion || '(empty - first run)'}`);

  const isFirstRun = !storedVersion || storedVersion.trim() === '';
  const versionChanged = currentVersion !== storedVersion;

  if (isFirstRun || versionChanged) {
    console.log(`Version change detected. First run: ${isFirstRun}, Changed: ${versionChanged}`);

    sendVersionUpdateEmail(currentVersion, isFirstRun);
    updateStoredVersion(currentVersion);
    console.log(`Updated stored version to: ${currentVersion}`);
  } else {
    console.log('Version unchanged. No email sent.');
  }

  console.log('Script finished');
}

/**
 * Attempts to trigger an API call that will return an error containing the API version
 * @return {Array<string>} Array of error messages from the API call
 */
function checkVersion() {
  const customerId = '123456789';
  const campaignId = '123456789';
  const negativeKeyword = 'test';

  const result = AdsApp.mutate({
    campaignOperation: {
      create: {
        negative: true,
        keyword: {
          text: negativeKeyword,
          matchType: 'EXACT'
        },
        campaign: `customers/${customerId}/campaigns/${campaignId}`,
      },
      "updateMask": "negative,keyword"
    }
  });

  return result.getErrorMessages();
}

/**
 * Extracts the API version (e.g., "v22") from an error message
 * Looks for pattern: google.ads.googleads.vXX.resources
 * @param {string} errorMessage - The error message containing the version
 * @return {string|null} The version string (e.g., "v22") or null if not found
 */
function extractVersionFromError(errorMessage) {
  if (!errorMessage) {
    console.warn('Error message is empty or null');
    return null;
  }

  console.log(`Extracting version from error message: ${errorMessage}`);

  // Pattern to match: google.ads.googleads.vXX.resources where XX is the version number
  const versionPattern = /google\.ads\.googleads\.(v\d+)\.resources/;
  const match = errorMessage.match(versionPattern);

  if (match && match[1]) {
    const version = match[1];
    console.log(`Successfully extracted version: ${version}`);
    return version;
  }

  console.warn(`Could not find version pattern in error message: ${errorMessage}`);
  return null;
}

/**
 * Gets the stored version from cell A1 of the first sheet in the spreadsheet
 * @return {string|null} The stored version string or null if cell is empty
 */
function getStoredVersion() {
  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sheet = spreadsheet.getSheets()[0];

    if (!sheet) {
      console.warn('No sheets found in spreadsheet');
      return null;
    }

    const versionCell = sheet.getRange(VERSION_CELL);
    const versionValue = versionCell.getValue();

    console.log(`Retrieved value from ${VERSION_CELL}: "${versionValue}"`);

    if (!versionValue || versionValue.toString().trim() === '') {
      return null;
    }

    return versionValue.toString().trim();
  } catch (error) {
    console.error(`Error reading stored version: ${error.toString()}`);
    return null;
  }
}

/**
 * Updates cell A1 of the first sheet with the new version
 * @param {string} version - The version string to store (e.g., "v22")
 */
function updateStoredVersion(version) {
  if (!version) {
    console.warn('Cannot update stored version: version is empty');
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    const sheet = spreadsheet.getSheets()[0];

    if (!sheet) {
      console.error('No sheets found in spreadsheet');
      return;
    }

    const versionCell = sheet.getRange(VERSION_CELL);
    versionCell.setValue(version);

    console.log(`Successfully updated ${VERSION_CELL} with version: ${version}`);
  } catch (error) {
    console.error(`Error updating stored version: ${error.toString()}`);
  }
}

/**
 * Sends an email notification about the version update
 * @param {string} version - The current API version (e.g., "v22")
 * @param {boolean} isFirstRun - Whether this is the first run of the script
 */
function sendVersionUpdateEmail(version, isFirstRun) {
  let emailBody;

  if (isFirstRun) {
    emailBody = `This is the first run of the Google Ads API Version Checker script.
    
The current Google Ads API for Scripts version is ${version}.

This email confirms that the script is working correctly. This may not indicate a new version upgrade - it's just the initial check.

See the release notes here: ${RELEASE_NOTES_URL}`;
  } else {
    emailBody = `The Google Ads API for Scripts has been upgraded to version ${version}.

See the release notes here: ${RELEASE_NOTES_URL}`;
  }

  try {
    MailApp.sendEmail({
      to: RECIPIENTS.join(','),
      subject: SUBJECT,
      body: emailBody
    });

    console.log(`Email sent successfully to: ${RECIPIENTS.join(', ')}`);
  } catch (error) {
    console.error(`Error sending email: ${error.toString()}`);
  }
}