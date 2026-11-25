# PMax Search Term N-Grams

Turn your Performance Max search terms into nGrams (1, 2, 3, and 4 word combinations) with aggregated metrics for analysis and negative keyword opportunities.

## Details

| | |
|---|---|
| **Category** | Performance Max |
| **Tags** | Performance Max, Search Terms, Analysis, Negative Keywords |
| **Difficulty** | Easy |
| **Schedule** | Weekly |
| **Makes Changes** | No |
| **Last Updated** | 2026-11-25 |

## Links

- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/pmax_search_term_n_grams/pmax_search_term_n_grams.js)

## Overview

This script will create 1, 2, 3, and 4 word nGrams from your Performance Max search terms.

**Important:** This script is for Performance Max campaigns only. See our Search Term N-Grams (Standard Shopping) script for Standard Shopping campaigns.

## What's an nGram?

Let's say we have two search terms:

- google ads api development
- bing ads api development

The **1 word nGrams** will be: google, bing, ads, api, development

The **2 word nGrams** will be: google ads, ads api, api development, bing ads

Importantly, metrics will also be combined to provide total clicks, impressions, conversions, conversion value etc. per nGram.

## Features

- 📊 **Multiple nGram sizes** - Creates 1, 2, 3, and 4 word nGrams
- 📈 **Aggregated metrics** - Impressions, clicks, conversions, conversion value, CTR
- 🎯 **Campaign filters** - Filter by campaign name (contains, not contains, equals)
- 📉 **Performance filters** - Set minimum impressions, clicks, or conversions
- 🗂️ **Segment by campaign** - Or aggregate at account level

## Configuration Options

| Setting | Description |
|---------|-------------|
| `LOOKBACK_DAYS` | Number of days to analyze (default: 30) |
| `SEGMENT_BY_CAMPAIGN` | Whether to segment by campaign or aggregate at account level |
| `CAMPAIGN_NAME_CONTAINS` | Filter campaigns containing this text |
| `MIN_IMPRESSIONS` | Minimum impressions filter |
| `MIN_CLICKS` | Minimum clicks filter |
| `MIN_CONVERSIONS` | Minimum conversions filter |

## Changes

**Version 3 (2026-11-25)**
- Uses the real Performance Max search terms (`campaign_search_term_view`) instead of Google's search term categories. If a campaign has none, or the report fails, it falls back to the categories as before.
- A search term that repeats a word (for example "shoes for shoes") now counts once towards that nGram, not twice.
- `LOOKBACK_DAYS` now covers exactly that many days, including today (it was one day more).
- `ENABLED_CAMPAIGNS_ONLY = false` now includes paused and removed campaigns (they were always left out).
- The header row is no longer included in the sort, and campaign IDs are no longer shown with thousands separators.
- At manager (MCC) level, a clear message is logged when `ACCOUNT_ID` is empty.

## What if I have suggestions?

Please let me know! Hearing your pain points is the number one way I can make improvements.
