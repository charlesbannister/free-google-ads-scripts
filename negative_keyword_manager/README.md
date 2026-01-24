# Negative Keyword Manager

Everything you need for reviewing search terms - ignore, add negatives, or add keywords.

## Details

| | |
|---|---|
| **Category** | Negative Keywords |
| **Tags** | Negative Keywords, Google Search |
| **Difficulty** | Easy |
| **Schedule** | Weekly |
| **Makes Changes** | Optional |
| **Last Updated** | 2026-01-23 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1MvCwzNCOIG3AO34b5yVe5VjcNFDwdvhCJdhrZQRXz58)
- [YouTube Tutorial](https://www.youtube.com/watch?v=CMdejdFU6Xg)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/negative_keyword_manager/negative_keyword_manager.js)

## Overview

This script is everything you need for reviewing search terms and deciding to:

- **Ignore them** - anything marked "ignored" won't appear in the reports again
- **Add Negative Keywords** - cut wasted spend by adding negative keywords
- **Add Keywords** - add your best performing search terms as keywords

## Adding New Rules

There are two steps to adding new rules:

1. Duplicate (copy and paste) an existing row and update its settings
2. Duplicate a sheet and name it accordingly

Note the "Output Sheet Name" needs to match the name of the sheet. Numbering the sheets is often easiest.

## Rules We Recommend

Every account is different, but as a general rule we recommend setting up the following rules. Remember these can be duplicated across multiple date ranges.

### Cream of the Crop / Best Performers

**Rule:** Strong CPA/ROAS AND High Clicks/Conversions/Conversion Value

**Reason:** Find potential keywords aka positive keywords.

### Zero Converters

**Rule:** Conversions < 0.01 AND High Clicks

**Reason:** Find potential negative keywords. Good option as an alert.

### Poor Performers (Converters)

**Rule:** Conversions > 0 AND High Clicks AND poor CPA/ROAS

**Reason:** Find potential negative keywords. Good option as an alert.

### Zero Clicks

**Rule:** Impressions > 100 AND Clicks < 1

**Reason:** Low CTR can be a sign of irrelevant search terms. Find potential negative keywords especially with a view to increase overall CTR.

### Poor CTR

**Rule:** Impressions > 100 AND Clicks > 0.99 AND CTR < 0.5%

**Reason:** Low CTR can be a sign of irrelevant search terms. Find potential negative keywords especially with a view to increase overall CTR.
