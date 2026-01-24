# Search Term Trends Report

Create multiple **account-level search term reports** with custom filters and 6 date ranges to compare.

## Details

| | |
|---|---|
| **Category** | Analysis |
| **Tags** | Search Terms, Google Search, Google Shopping, Analysis |
| **Difficulty** | Easy |
| **Schedule** | Daily |
| **Makes Changes** | No |
| **Last Updated** | 2024-02 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1RQn6LN1H8shjKbZqbGpgmxPVjwWe-y4Pwf7ujIbWnM4)
- [YouTube Tutorial](https://www.youtube.com/watch?v=5rztOybG7vk)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/main/scripts/search-term-trends.js)

## What This Script Does

Search term stats as an **average across 6 different time periods**.

**As an average?** E.g. if the lookback window (days) is 365, it's the total clicks across 365 days (to yesterday) divided by 365.

## Recommended Schedule

- **Hourly** if looking at today's data
- **Daily** otherwise

## Adding More Reports

To create a new report, just **duplicate an existing report tab** then add your settings.

### How many reports can I add?

The only limit is the script runtime of 30 minutes and the capacity of Google Sheets (10M cells max, but they can become frustratingly slow before that point).

**Should the script timeout**, either reduce the number of reports or add further Metric or Campaign filters.

💡 You can install the same script multiple times—just use a different Sheet URL each time.

## Have Suggestions?

Please let me know! Hearing your pain points is the number one way I can make improvements.
