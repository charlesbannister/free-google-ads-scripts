# Placement Exclusions - Semi-automated (PMax, Display, YouTube)

Identify unwanted placements across Performance Max, Display, and YouTube campaigns. Review in Google Sheets and exclude with checkboxes.

## Details

| | |
|---|---|
| **Category** | Placement Exclusions |
| **Tags** | Performance Max, Display, YouTube, Placements, Exclusions |
| **Difficulty** | Intermediate |
| **Schedule** | Weekly |
| **Makes Changes** | Optional |
| **Last Updated** | 2026-01-20 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1jG_igH1QGdyBSbeqj2ELxZEg3uOFYd3wcWaQ_eDfY9o)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/placement_exclusions/pmax_placement_exclusions.js)

## Overview

This script helps you identify unwanted placements, which you can optionally exclude via checkboxes in a Google Sheet.

**Works across Performance Max, Display, and YouTube campaigns** - all from one script!

## What This Script Does

- 📊 **Generates placement reports** - Pulls placement data from PMax, Display, and YouTube campaigns
- ✅ **Checkbox-based exclusions** - Review placements in Google Sheets and tick boxes to exclude
- 🤖 **Optional ChatGPT integration** - Analyze website content to help identify irrelevant placements
- 📝 **Exclusion logging** - Maintains a log of what's been excluded and when

## How It Works

1. The script pulls placement data from your eligible campaigns
2. Data is written to separate sheets by placement type (Website, YouTube, Mobile App, Google Products)
3. Review the placements and tick the checkbox for any you want to exclude
4. Run the script again to apply your exclusions

## Campaigns Supported

- ✅ Performance Max campaigns
- ✅ Display campaigns
- ✅ YouTube/Video campaigns

## Configuration Options

| Setting | Description |
|---------|-------------|
| `Shared Exclusion List Name` | Name of the shared exclusion list to add placements to |
| `Lookback Window (Days)` | Number of days of data to analyze |
| `Minimum Impressions` | Only show placements with impressions above this |
| `Minimum Clicks` | Only show placements with clicks above this |
| `Minimum Cost` | Only show placements with cost above this |
| `Maximum Conversions` | Only show placements with conversions below this |
| `Enabled Campaigns Only` | Whether to only include enabled campaigns |

## Features

- **Placement type sheets** - Separate output sheets for Website, YouTube, Mobile App, and Google Products
- **Automated mode** - Optionally pre-tick checkboxes based on your rules
- **Campaign name filters** - Filter by campaign name contains/not contains
- **Placement filters** - Filter by display name, target URL, and TLDs
- **ChatGPT integration** - Optional AI analysis of website content

## What if I have suggestions?

Please let me know! Hearing your pain points is the number one way I can make improvements.
