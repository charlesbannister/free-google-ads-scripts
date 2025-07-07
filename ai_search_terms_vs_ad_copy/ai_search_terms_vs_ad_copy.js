/**
 * Search Term Vs Ad Copy AI Analysis
 * This script uses AI (ChatGPT) to compare search terms to ad copy,
 * creating a relevance score and writing it to a Google Sheet.
 * @author Charles Bannister (shabba.io)
 * @version 1.0.0
 * Written as part of a Vibe Coding Google Ads Script Tutorial
 * Tutorial: https://www.youtube.com/watch?v=ETN5HpFbvZY
 */

// Query builder: https://developers.google.com/google-ads/api/fields/v16/ad_group_ad_asset_view_query_builder

//Template Sheet (click to make a copy): https://docs.google.com/spreadsheets/d/1yEBBiX4w4-RCTICtDy-YBoe-aVghDOJc0uWLmCsuEX4/copy
const SPREADSHEET_URL = 'YOUR_SPREADSHEET_URL_HERE';

const DEBUG_MODE = true; // Set to false to reduce logging

const LOOKBACK_WINDOW_DAYS = 30; // Look back window for query date range
const MINIMUM_CLICKS = 1; // Minimum clicks to include a search term
const OPENAI_SETTINGS_SHEET = 'settings'; // Sheet name containing OpenAI API key
const OUTPUT_SHEET_NAME = 'results'; // Sheet where results are written

const OPENAI_SYSTEM_ROLE = 'You are an expert Google Ads evaluator. You write short, pithy explanations with just the core info and nothing more.'; // Used as the system role to guide ChatGPT's perspective
const OPENAI_PROMPT_TEMPLATE = `
  I will provide a search term and its parent ad copy (headlines and descriptions).
  Rate the search term's relevance to the ad copy on a 0-10 scale.
  Briefly justify your answer.
  Search Term: "{{searchTerm}}"\n\nAd Headline(s): {{headlines}}\n\nAd Description(s): {{descriptions}}`;
// Template for generating the prompt sent to ChatGPT. Fields in {{brackets}} will be dynamically replaced.

function main() {
	console.log(`Script started`);

	const searchTermData = getSearchTermData();
	const adCopyData = getResponsiveSearchAds();
	const prompts = buildRelevancePrompts(searchTermData, adCopyData);
	const apiKey = getOpenAiApiKey();
	const responses = callChatGptApi(prompts, apiKey);

	writeResponsesToSheet(responses);

	console.log(`\nChatGPT Responses (first 3):`);
	responses.slice(0, 3).forEach((res, index) => {
		console.log(`\nResponse ${index + 1}:`);
		console.log(`Search Term: ${res.searchTerm}`);
		console.log(`Ad ID: ${res.adId}`);
		console.log(`Score: ${res.relevanceScore}`);
		console.log(`Explanation: ${res.explanation}`);
	});

	console.log(`\nTotal responses: ${responses.length}`);
	console.log(`Script finished`);
}

function getSearchTermData() {
	const query = getSearchTermGaqlQuery();
	const report = getSearchTermReport(query);
	const flatRows = extractFlatSearchTermRows(report);
	return flatRows.filter(row => row.clicks > MINIMUM_CLICKS);
}

function getSearchTermGaqlQuery() {
	const endDate = getGoogleAdsApiFormattedDate(0);
	const startDate = getGoogleAdsApiFormattedDate(LOOKBACK_WINDOW_DAYS);

	const query = `
    SELECT
      campaign.id,
      campaign.name,
      ad_group.id,
      ad_group.name,
      search_term_view.search_term,
      metrics.clicks,
      metrics.impressions,
      metrics.conversions,
      metrics.conversions_value,
      metrics.cost_micros
    FROM search_term_view
    WHERE
      campaign.advertising_channel_type = 'SEARCH'
      AND segments.date BETWEEN '${startDate}' AND '${endDate}'
    ORDER BY metrics.clicks DESC
  `;

	console.log(`GAQL Query Used:\n${query}`);
	return query;
}

function getSearchTermReport(query) {
	try {
		const report = AdsApp.report(query);
		return report;
	} catch (error) {
		console.error(`Error fetching search term report: ${error.message}`);
		console.error(`Validate your query here: https://developers.google.com/google-ads/api/fields/v16/search_term_view_query_builder`);
		throw error;
	}
}

function extractFlatSearchTermRows(report) {
	const rows = report.rows();
	const data = [];

	while (rows.hasNext()) {
		const row = rows.next();

		data.push({
			campaignId: row['campaign.id'],
			campaignName: row['campaign.name'],
			adGroupId: row['ad_group.id'],
			adGroupName: row['ad_group.name'],
			searchTerm: row['search_term_view.search_term'],
			clicks: parseInt(row['metrics.clicks'], 10),
			impressions: parseInt(row['metrics.impressions'], 10),
			conversions: parseFloat(row['metrics.conversions'] || 0),
			conversionsValue: parseFloat(row['metrics.conversions_value'] || 0),
			cost: parseFloat(row['metrics.cost_micros']) / 1000000,
		});
	}

	return data;
}

function getResponsiveSearchAds() {
	const query = `
    SELECT
      ad_group.id,
      ad_group.name,
      ad_group_ad.ad.id,
      ad_group_ad_asset_view.field_type,
      asset.text_asset.text
    FROM ad_group_ad_asset_view
    WHERE
      ad_group_ad_asset_view.enabled = true
      AND asset.type = 'TEXT'
  `;

	console.log(`GAQL Query Used for RSAs:\n${query}`);

	const report = AdsApp.report(query);
	const rows = report.rows();
	const grouped = {};

	while (rows.hasNext()) {
		const row = rows.next();
		const adGroupId = row['ad_group.id'];
		const adGroupName = row['ad_group.name'];
		const adId = row['ad_group_ad.ad.id'];
		const fieldType = row['ad_group_ad_asset_view.field_type'];
		const text = row['asset.text_asset.text'];

		const key = `${adGroupId}_${adId}`;
		if (!grouped[key]) {
			grouped[key] = {
				adGroupId,
				adGroupName,
				adId,
				headlines: [],
				descriptions: []
			};
		}

		if (fieldType === 'HEADLINE') {
			grouped[key].headlines.push(text);
		} else if (fieldType === 'DESCRIPTION') {
			grouped[key].descriptions.push(text);
		}
	}

	return Object.values(grouped).map(ad => ({
		adGroupId: ad.adGroupId,
		adGroupName: ad.adGroupName,
		adId: ad.adId,
		headlines: ad.headlines.join(' | '),
		descriptions: ad.descriptions.join(' | ')
	}));
}

function buildRelevancePrompts(searchTerms, rsaAds) {
	const prompts = [];

	searchTerms.forEach(term => {
		const matchingAds = rsaAds.filter(ad => ad.adGroupId === term.adGroupId);

		matchingAds.forEach(ad => {
			const prompt = OPENAI_PROMPT_TEMPLATE
				.replace('{{searchTerm}}', term.searchTerm)
				.replace('{{headlines}}', ad.headlines)
				.replace('{{descriptions}}', ad.descriptions);

			prompts.push({
				searchTerm: term.searchTerm,
				clicks: term.clicks,
				campaignName: term.campaignName,
				adGroupName: term.adGroupName,
				adGroupId: term.adGroupId,
				adId: ad.adId,
				headlines: ad.headlines,
				descriptions: ad.descriptions,
				prompt
			});
		});
	});

	return prompts;
}

function getOpenAiApiKey() {
	const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
	const sheet = spreadsheet.getSheetByName(OPENAI_SETTINGS_SHEET);
	if (!sheet) throw new Error(`Sheet '${OPENAI_SETTINGS_SHEET}' not found.`);

	const apiKey = sheet.getRange('B1').getValue();
	if (!apiKey || apiKey === 'PASTE_YOUR_OPENAI_API_KEY_HERE') {
		throw new Error('OpenAI API key is missing or not set in settings sheet cell B1.');
	}

	return apiKey;
}

function callChatGptApi(prompts, apiKey) {
	const responses = [];

	prompts.forEach(promptObj => {
		const payload = {
			model: 'gpt-3.5-turbo',
			messages: [
				{ role: 'system', content: OPENAI_SYSTEM_ROLE },
				{ role: 'user', content: promptObj.prompt }
			],
			temperature: 0.2
		};

		const options = {
			method: 'post',
			contentType: 'application/json',
			headers: { Authorization: `Bearer ${apiKey}` },
			payload: JSON.stringify(payload),
			muteHttpExceptions: true
		};

		try {
			const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
			const statusCode = response.getResponseCode();

			if (statusCode !== 200) {
				console.error(`OpenAI API returned HTTP ${statusCode}`);
				console.error(response.getContentText());
				return;
			}
			let json;
			try {
				json = JSON.parse(response.getContentText());
			} catch (parseError) {
				console.error('Failed to parse OpenAI response as JSON');
				console.error(response.getContentText());
				return;
			}
			const content = json.choices[0].message.content;

			const match = content.match(/(\d+)/);
			const score = match ? parseInt(match[1], 10) : null;

			responses.push({
				searchTerm: promptObj.searchTerm,
				clicks: promptObj.clicks,
				campaignName: promptObj.campaignName,
				adGroupName: promptObj.adGroupName,
				adGroupId: promptObj.adGroupId,
				adId: promptObj.adId,
				headlines: promptObj.headlines,
				descriptions: promptObj.descriptions,
				relevanceScore: score,
				explanation: content
			});
		} catch (e) {
			console.error(`Error calling OpenAI for prompt: ${promptObj.prompt}`);
			console.error(e.message);
		}
	});

	return responses;
}

function writeResponsesToSheet(responses) {
	const spreadsheet = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
	let sheet = spreadsheet.getSheetByName(OUTPUT_SHEET_NAME);

	if (!sheet) {
		sheet = spreadsheet.insertSheet(OUTPUT_SHEET_NAME);
	} else {
		sheet.clear();
	}

	const headers = [
		'Campaign Name', 'Ad Group Name', 'Search Term', 'Clicks', 'Ad Group ID', 'Ad ID',
		'Headlines', 'Descriptions', 'Relevance Score', 'Explanation'
	];

	const data = responses.map(row => [
		row.campaignName,
		row.adGroupName,
		row.searchTerm,
		row.clicks,
		row.adGroupId,
		row.adId,
		row.headlines,
		row.descriptions,
		row.relevanceScore,
		row.explanation
	]);

	sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
	if (data.length) {
		sheet.getRange(2, 1, data.length, headers.length).setValues(data);
	}
	sheet.setFrozenRows(1);
	sheet.autoResizeColumns(1, headers.length);
}

function getGoogleAdsApiFormattedDate(daysAgo = 0) {
	const date = new Date();
	date.setDate(date.getDate() - daysAgo);
	const year = date.getFullYear();
	const month = String(date.getMonth() + 1).padStart(2, '0');
	const day = String(date.getDate()).padStart(2, '0');
	return `${year}-${month}-${day}`;
}
