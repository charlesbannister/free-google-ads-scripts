# AI Search Terms Vs Ad Copy

Use ChatGPT to score how relevant your search terms are to your ad copy, with AI-generated explanations.

## Details

| | |
|---|---|
| **Category** | Analysis |
| **Tags** | AI, Google Search, Negative Keywords, Analysis, Reporting, Ad Copy, User Journey |
| **Difficulty** | Intermediate |
| **Schedule** | Weekly |
| **Makes Changes** | No |
| **Last Updated** | 2026-01-23 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1yEBBiX4w4-RCTICtDy-YBoe-aVghDOJc0uWLmCsuEX4)
- [YouTube Tutorial](https://www.youtube.com/watch?v=ETN5HpFbvZY)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/ai_search_terms_vs_ad_copy/ai_search_terms_vs_ad_copy.js)

## What It Does

This script uses **ChatGPT** to analyze how well your search terms match your ad copy. For each search term, it compares the query to your ad headlines and descriptions, then provides a **relevance score (0-10)** with a brief explanation.

1. Pulls your search terms from the last 30 days
2. Gets all your responsive search ad copy (headlines & descriptions)
3. Sends each search term + ad copy pair to ChatGPT
4. Writes results to a Google Sheet with scores and explanations

## Requirements

- **OpenAI API Key** - You'll need a ChatGPT API key (stored in the sheet's "settings" tab)
- **Google Ads search campaigns** with responsive search ads
- **At least some search term data** to analyze

## Output

Each row in the results sheet includes:
- Campaign & Ad Group names
- Search term and click count
- Ad headlines and descriptions
- **Relevance Score** (0-10)
- **AI Explanation** of why the score was given

## Use Cases

- Find search terms that don't match your ad messaging
- Identify opportunities to improve ad relevance
- Spot potential negative keyword candidates (low relevance scores)

**Note:** This script was created as part of a [Vibe Coding tutorial](https://www.youtube.com/watch?v=ETN5HpFbvZY) demonstrating how to build AI-powered Google Ads scripts.
