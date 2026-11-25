# Search Term N-Grams (Standard Shopping)

Turn your search terms into nGrams. Use as a reporting tool, alert system, or add negative keywords directly to a Negative Keyword List.

## Details

| | |
|---|---|
| **Category** | Shopping |
| **Tags** | Search Terms, Google Shopping, Analysis, Negative Keywords, Reporting, Alerts |
| **Difficulty** | Intermediate |
| **Schedule** | Weekly |
| **Makes Changes** | Optional |
| **Last Updated** | 2026-01-20 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1L6ty0u7OtD3Ed5h4SBN4mCwkbAqYsaw_18kOLQDr6Tk)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/search_term_n_grams_standard/search_term_n_grams.js)

## Turn your search terms into nGrams

This script will create 1, 2, 3, and 4 word nGrams from search terms.

**Important:** This script is for Standard Shopping campaigns only. It will not work for Performance Max. See our PMax Search Term N-Grams script for Performance Max campaigns.

## What can this script do?

- 📊 **Reporting** - Generate detailed nGram reports in Google Sheets
- 🚨 **Alerts** - Send an email if an nGram gets over X clicks without a conversion
- ➖ **Add Negative Keywords** - Automatically add nGrams as negative keywords to a Negative Keyword List (optional)

## What's an nGram?

Let's say we have two search terms:

- google ads api development
- bing ads api development

The **1 word nGrams** will be: google, bing, ads, api, development

The **2 word nGrams** will be: google ads, ads api, api development, bing ads

Importantly, metrics will also be combined to provide total clicks, cost, conversions etc. per nGram.

## Use Cases

Use this script to:

- Run a report to analyze search term patterns
- Set up alerts e.g. send an email if an nGram gets over 100 clicks without a conversion
- Automatically add poor-performing nGrams as negative keywords

## Changes

**Version 1.1.0 (2026-11-25)**
- A search term that repeats a word (for example "shoes for shoes") now counts once towards that nGram, not twice.
- Extra spaces in a search term no longer create blank nGrams.
- ROAS is 0 instead of Infinity for an nGram with conversion value but no cost.
- Output tabs grow when needed, so more than 1,000 nGrams, or a newly created output sheet, no longer fails with "the coordinates of the range are outside the dimensions of the sheet". The formatting ranges now match the data.
- A number in a comma-separated setting (for example a campaign name filter of `2024`) no longer stops the script.

## What if I have suggestions?

Please let me know! Hearing your pain points is the number one way I can make improvements.
