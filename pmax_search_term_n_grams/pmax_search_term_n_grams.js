/**
 * Performance Max Search Term nGrams
 * @author Charles Bannister (https://www.linkedin.com/in/charles-bannister/)
 * Non-PMax version: https://shabba.io/script/4
 */

//You may also be interested in this Chrome keyword wrapper!
//https://chrome.google.com/webstore/detail/keyword-wrapper/paaonoglkfolneaaopehamdadalgehbb

const VERSION = 2;

// START OF OPTIONS

const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';
// Create your Google Sheet by typing sheets.new in the browser
// Copy the URL from the browser and paste it here (between the quotes)

const CURRENCY_SYMBOL = "$";
// For formatting monetary values

const LOOKBACK_DAYS = 30;
// Includes today

const CAMPAIGN_NAME_CONTAINS = "";
// Set to "" disable

const CAMPAIGN_NAME_NOT_CONTAINS = "";
// Set to "" disable

const CAMPAIGN_NAME_EQUALS = "";
// Set to "" disable

const ENABLED_CAMPAIGNS_ONLY = true;
// Whether to only include enabled campaigns
// Paused + removed will be included if false

const SEGMENT_BY_CAMPAIGN = false;
// Whether to segment by campaign
// Campaign will be "All" if false (account level nGrams)

/* ---- Search Term Filters ---- */
// Note these filters are applied when initially
// fetching the PMax data
// They may be necessary (especially for larger accounts)
// but could result in innacurate nGrams

const MIN_IMPRESSIONS = 0;
// Impressions > this number

const MIN_CLICKS = 0;
// Clicks > this number

const MIN_CONVERSIONS = 0;
// Conversions > this number

// Feel free to edit these filters directly too
// ChatGPT, etc. can be helpful here
// You can also filter by:
// metrics.cost
// metrics.conversions_value
const performanceFilters = [
    {
        field: 'metrics.impressions',
        operator: '>',
        value: MIN_IMPRESSIONS
    },
    {
        field: 'metrics.clicks',
        operator: '>',
        value: MIN_CLICKS
    },
    {
        field: 'metrics.conversions',
        operator: '>',
        value: MIN_CONVERSIONS
    }
]

const USE_DUMMY_DATA = false;
//For testing. Should be false.

const ACCOUNT_ID = '';
//This is for running a single account at MCC level
//Add the ID (000-000-0000) of the account you want to run

function main() {
    console.log('Started');
    if (USE_DUMMY_DATA) {
        console.warn("WARNING: Using dummy data. Set USE_DUMMY_DATA to false to use real data.");
    }
    if (!isMCC()) {
        runAccount();
        return;
    }
    MccApp.accounts()
        .withIds([ACCOUNT_ID])
        .withLimit(50)
        .executeInParallel("runAccount");
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
 * Validates and returns the spreadsheet object
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null} The spreadsheet object or null if invalid
 */
function validateSpreadsheet() {
    if (SPREADSHEET_URL === "YOUR_SPREADSHEET_URL_HERE") {
        console.log("ERROR: Please replace 'YOUR_SPREADSHEET_URL_HERE' with your actual Google Sheet URL in the script.");
        return null;
    }

    try {
        const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
        if (!spreadsheet) {
            console.log("ERROR: Could not access spreadsheet. Please check URL and permissions.");
            return null;
        }
        console.log("Successfully accessed spreadsheet");
        return spreadsheet;
    } catch (e) {
        console.log(`ERROR: Failed to access spreadsheet: ${e.message}`);
        return null;
    }
}

/**
 * Fetches PMAX data based on configuration
 * @returns {Array} Array of PMAX data rows
 */
function fetchPmaxData() {
    const lookbackDays = LOOKBACK_DAYS;
    console.log(`Using lookback days: ${lookbackDays}`);

    const entityConditions = getEntityConditions();
    console.log(`Entity conditions: ${JSON.stringify(entityConditions)}`);

    return _fetchPmaxData_(lookbackDays, entityConditions);
}

/**
 * Main function executed by Google Ads Scripts.
 */
function runAccount() {
    try {
        console.log('Starting runAccount function');

        const spreadsheet = validateSpreadsheet();
        if (!spreadsheet) {
            return;
        }

        const pmaxRows = USE_DUMMY_DATA ? createDummyPmaxData() : fetchPmaxData();
        console.log(`Retrieved ${pmaxRows.length} PMAX rows`);
        if (pmaxRows.length === 0) {
            console.log('No data to write to sheet');
            return;
        }

        //ngrams
        const ngramsBySize = generateNgrams(pmaxRows);

        for (const [sheetName, ngramRows] of Object.entries(ngramsBySize)) {
            const sheet = _getOrCreateSheet_(spreadsheet, sheetName);
            writeToSheet(sheet, ngramRows);
            console.log(`Successfully wrote ${sheetName} data to sheet`);
        }
    } catch (e) {
        console.log(`Error running account: ${e.message}`);
        console.log(`Stack trace: ${e.stack}`);
    }
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

function writeToSheet(sheet, pmaxRows, startRow = 1) {
    sheet.clear();
    if (pmaxRows.length === 0) {
        console.log('No data to write to sheet');
        sheet.getRange(1, 1, 1, 1).setValue('No data to write to sheet. Check the filters and try again.');
        return;
    }

    // Convert objects to array of arrays with headers
    const headers = [
        'nGram', 'Count', 'Campaign ID', 'Campaign Name',
        'Impressions', 'Clicks', 'Conversions', 'Conversion Value',
        'CTR'
    ];
    const data = pmaxRows.map(row => [
        row.searchTerm,
        row.ngramCount,
        row.campaignId,
        row.campaignName,
        row.impressions,
        row.clicks,
        row.conversions,
        row.conversions_value,
        row.ctr,

    ]);

    // Add headers to the data
    const formattedData = [headers, ...data];

    const numRows = formattedData.length;
    const numCols = formattedData[0].length;
    console.log(`Writing ${numRows} data rows (${numCols} columns) to sheet: "${sheet.getName()}" starting at row ${startRow}`);

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
    dataRange.setValues(formattedData);

    const sortByColumns = [];
    sortByColumns.push({ column: 7, ascending: false });
    sheet.getRange(startRow, 1, numRows, numCols).sort(sortByColumns);

    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);

    var currencyFormat = CURRENCY_SYMBOL + "#,##0.00";

    const formatting = ["#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", currencyFormat, "0.00%"];
    const validFormatting = formatting.length === headers.length;
    if (!validFormatting) {
        console.error("Formatting array length does not match headers length");
        return;
    }

    for (let i = 0; i < headers.length; i++) {
        sheet.getRange(1, i + 1, numRows, 1).setNumberFormat(formatting[i]);
    }
}

function getEntityConditions() {
    const entityConditions = [];
    if (CAMPAIGN_NAME_CONTAINS) {
        entityConditions.push({ field: 'campaign.name', operator: 'contains', value: CAMPAIGN_NAME_CONTAINS });
    }
    if (CAMPAIGN_NAME_NOT_CONTAINS) {
        entityConditions.push({ field: 'campaign.name', operator: 'not_contains', value: CAMPAIGN_NAME_NOT_CONTAINS });
    }
    if (CAMPAIGN_NAME_EQUALS) {
        entityConditions.push({ field: 'campaign.name', operator: 'equals', value: CAMPAIGN_NAME_EQUALS });
    }
    if (ENABLED_CAMPAIGNS_ONLY) {
        entityConditions.push({ field: 'campaign.status', operator: 'equals', value: 'ENABLED' });
    }
    return entityConditions;
}

/**
 * Fetches Performance Max search terms data for a rule
 * @param {number} lookbackDays Number of days to look back
 * @param {Array} entityConditions Conditions from the rule
 * @returns {Array} Array of PMAX search term rows or empty array if none found
 * @private
 */
function _fetchPmaxData_(lookbackDays, entityConditions) {
    let pmaxRows = [];

    try {
        // Create PMAX report builder
        const pmaxReportBuilder = new PMaxReportBuilder(lookbackDays, entityConditions);

        // First check if any PMAX campaigns match the conditions
        const matchingCampaigns = pmaxReportBuilder.getMatchingPmaxCampaigns();
        console.log(`matchingCampaigns: ${JSON.stringify(matchingCampaigns)}`);

        if (matchingCampaigns && matchingCampaigns.length > 0) {
            console.log(`Found ${matchingCampaigns.length} Performance Max campaigns. Fetching search terms...`);

            // Get search terms for the matching campaigns
            const campaignPmaxRows = pmaxReportBuilder.getPmaxSearchTerms(matchingCampaigns);
            pmaxRows = [...pmaxRows, ...campaignPmaxRows];
            console.log(`pmaxRows length: ${pmaxRows.length}`);
            if (pmaxRows && pmaxRows.length > 0) {
                console.log(`Found ${pmaxRows.length} PMAX search terms`);
            } else {
                console.log(`No PMAX search terms found`);
            }
        }
    } catch (e) {
        console.log(`ERROR executing PMAX report: ${e.message}.`);
    }
    console.log(`Returning ${pmaxRows.length} PMAX rows`);
    return pmaxRows;
}

/**
 * Gets a date string formatted as YYYY-MM-DD for GAQL.
 * @param {number} daysAgo Number of days ago from today (0 for today).
 * @return {string} Formatted date string.
 * @private
 */
function getFormattedDateString_(daysAgo) {
    const date = new Date();
    date.setDate(date.getDate() - daysAgo);
    return Utilities.formatDate(date, AdsApp.currentAccount().getTimeZone(), 'yyyy-MM-dd');
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

                let performanceConditions = [];
                for (const filter of performanceFilters) {
                    performanceConditions.push(`${filter.field} ${filter.operator} ${filter.value}`);
                }


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
                    AND ${performanceConditions.join(' AND ')}
                `;

                console.log(`Executing PMAX search terms query for campaign ${campaignId} (${campaignName})`);
                console.log(`Query: ${query}`);
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
 * Generates the GAQL date range condition string.
 * @param {number} lookbackDays Number of days to look back.
 * @return {string} The GAQL condition string (e.g., "segments.date BETWEEN 'YYYY-MM-DD' AND 'YYYY-MM-DD'").
 * @private
 */
function _getDateRangeCondition_(lookbackDays) {
    const endDate = getFormattedDateString_(0); // Today
    const startDate = getFormattedDateString_(lookbackDays);
    return `segments.date BETWEEN '${startDate}' AND '${endDate}'`;
}

/**
 * Extracts n-grams from a single search term
 * @param {string} searchTerm The search term to process
 * @param {number} n The size of n-grams to generate
 * @returns {Array<string>} Array of n-grams
 */
function extractNgramsFromTerm(searchTerm, n) {
    const words = searchTerm.toLowerCase().trim().split(/\s+/);
    const ngrams = [];
    for (let i = 0; i <= words.length - n; i++) {
        ngrams.push(words.slice(i, i + n).join(' '));
    }
    return ngrams;
}

/**
 * Creates n-gram objects with original metrics
 * @param {Object} row Original search term data
 * @param {string} ngram The n-gram text
 * @param {number} ngramSize The size of the n-gram
 * @returns {Object} N-gram object with metrics
 */
function createNgramObject(row, ngram, ngramSize) {
    return {
        searchTerm: ngram,
        ngramSize: ngramSize,
        campaignId: row.campaignId,
        campaignName: row.campaignName,
        impressions: row.impressions,
        clicks: row.clicks,
        cost: row.cost,
        conversions: row.conversions,
        conversions_value: row.conversions_value
    };
}

/**
 * Aggregates metrics for a group of n-grams
 * @param {Array} ngrams Array of n-gram objects
 * @returns {Object} Aggregated metrics by n-gram
 */
function aggregateNgramMetrics(ngrams) {
    const aggregated = {};
    for (const ngram of ngrams) {
        let key;
        if (SEGMENT_BY_CAMPAIGN) {
            key = `${ngram.searchTerm}_${ngram.ngramSize}_${ngram.campaignId}`;
        } else {
            key = `${ngram.searchTerm}_${ngram.ngramSize}`;
        }
        if (!aggregated[key]) {
            aggregated[key] = {
                searchTerm: ngram.searchTerm,
                ngramSize: ngram.ngramSize,
                ngramCount: 0,
                campaignId: SEGMENT_BY_CAMPAIGN ? ngram.campaignId : 'All',
                campaignName: SEGMENT_BY_CAMPAIGN ? ngram.campaignName : 'All',
                impressions: 0,
                clicks: 0,
                conversions: 0,
                conversions_value: 0
            };
        }

        aggregated[key].ngramCount += 1;
        aggregated[key].impressions += ngram.impressions;
        aggregated[key].clicks += ngram.clicks;
        aggregated[key].conversions += ngram.conversions;
        aggregated[key].conversions_value += ngram.conversions_value;

        // Calculate derived metrics
        aggregated[key].ctr = aggregated[key].impressions > 0 ? aggregated[key].clicks / aggregated[key].impressions : 0;
    }
    return Object.values(aggregated);
}

/**
 * Generates and aggregates n-grams from search terms
 * @param {Array} rows Array of search term data rows
 * @returns {Object} Object containing n-grams by size
 */
function generateNgrams(rows) {
    const ngramsBySize = {
        '1-grams': [],
        '2-grams': [],
        '3-grams': [],
        '4-grams': []
    };

    // Generate all n-grams
    const allNgrams = [];
    for (const row of rows) {
        for (let n = 1; n <= 4; n++) {
            const ngrams = extractNgramsFromTerm(row.searchTerm, n);
            for (const ngram of ngrams) {
                allNgrams.push(createNgramObject(row, ngram, n));
            }
        }
    }

    // Aggregate and organize by size
    const aggregatedNgrams = aggregateNgramMetrics(allNgrams);
    for (const ngram of aggregatedNgrams) {
        const sizeKey = `${ngram.ngramSize}-grams`;
        ngramsBySize[sizeKey].push(ngram);
    }

    return ngramsBySize;
}

/**
 * Creates dummy PMAX data for testing
 * @returns {Array} Array of dummy PMAX rows
 */
function createDummyPmaxData() {
    return [
        {
            searchTerm: "blue running shoes nike",
            campaignId: "123",
            campaignName: "Test Campaign 1",
            impressions: 1006,
            clicks: 52,
            cost: 100.21,
            conversions: 5.2,
            conversions_value: 500.21
        },
        {
            searchTerm: "red nike running shoes",
            campaignId: "123",
            campaignName: "Test Campaign 1",
            impressions: 807,
            clicks: 41,
            cost: 80.7,
            conversions: 4.1,
            conversions_value: 400.7
        },
        {
            searchTerm: "nike shoes black",
            campaignId: "456",
            campaignName: "Test Campaign 2",
            impressions: 607,
            clicks: 31,
            cost: 60.7,
            conversions: 3.1,
            conversions_value: 300.7
        }
    ];
}