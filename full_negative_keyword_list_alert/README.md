# Negative Keyword List Full Checker

Get alerted when your shared negative keyword lists hit the 5,000 keyword limit. One alert per list, tracked via Google Sheets.

## Details

| | |
|---|---|
| **Category** | Negative Keywords |
| **Tags** | Negative Keywords, Google Search |
| **Difficulty** | Easy |
| **Schedule** | Daily |
| **Makes Changes** | No |
| **Last Updated** | 2026-01-23 |

## Links

- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/full_negative_keyword_list_alert/full_negative_keyword_list_alert.js)

## Overview

Google Ads shared negative keyword lists have a **5,000 keyword limit**. When a list hits capacity, you can't add more negatives—and you might not even notice until wasted spend piles up.

This script monitors your shared negative keyword lists and sends you an **email alert when any list reaches the limit**. Each list only triggers one alert (tracked in a Google Sheet), so you won't get spammed.

## How It Works

1. Scans all your shared negative keyword lists
2. Checks keyword counts against the 5,000 limit
3. Sends an email alert for any full lists
4. Records alerted lists in a Google Sheet (no repeat alerts)

## Setup

1. Create a new Google Sheet (type `sheets.new` in your browser)
2. Add your sheet URL and email to the script config
3. Authorize and schedule to run daily

## Why You Need This

- **Full lists = no protection** – New irrelevant terms slip through
- **Easy to miss** – Google doesn't warn you when lists are full
- **One-time alert** – No alert fatigue, just actionable notifications

**Note:** If a list fills up, gets cleaned out, then fills up again, it won't trigger another alert unless you remove its ID from the tracking sheet.
