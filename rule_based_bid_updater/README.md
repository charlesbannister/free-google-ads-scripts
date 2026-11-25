# Rule Based Bid Updater / Pauser (Keywords & Products)

Setup rules to update keyword and product bids based on cost, conversions, ROAS, CPA, and more. Includes pause functionality.

## Details

| | |
|---|---|
| **Category** | Bidding |
| **Tags** | Bidding, Keywords, Products, Google Shopping, Optimization |
| **Difficulty** | Intermediate |
| **Schedule** | Daily |
| **Makes Changes** | Yes |
| **Last Updated** | 2026-01-24 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1RjDClzOKNoe7_5JZwCudllh4objhE73B-NAaTpGHMxc)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/rule_based_bid_updater/rule_based_bid_updater.js)

## Overview

Setup rules to update keyword and product bids based on cost, conversions, ROAS, CPA, and more. This script also includes pause functionality.

## When to Use Manual Bidding

I've been writing custom scripts for nearly a decade now and have noticed requests for bidding scripts have slowed.

That's a good thing. We should all be leaning into Automated Bidding where possible.

However, I still feel manual bidding can be a valuable tool in a PPC manager's arsenal - usually when there isn't enough data for automated bidding.

It can also be wise to outright exclude a product or pause a keyword that's gone awry.

## Example Rules

| Rule | Action |
|------|--------|
| No conversions and > 200 clicks | Pause the keyword |
| No conversions and > 100 clicks | Lower the bid by 10% |
| ROAS above target and > 50 clicks | Increase the bid by 10% |
| < 10 impressions | Increase the bid by 10% |

## Features

- **Keyword bid adjustments** - Increase or decrease bids based on performance
- **Product bid adjustments** - Works with Google Shopping campaigns
- **Pause functionality** - Automatically pause underperformers
- **Multiple metrics** - Filter by cost, conversions, ROAS, CPA, clicks, impressions
- **Customizable thresholds** - Set your own rules in the Google Sheet

## Coming Soon

I'm working on a pro version with extra options and DSA support - let me know if you're interested and I'll bump it up the queue!

## Changes

- **1.0.8 (2026-11-25)**: fixed the date range (quoted dates, the full N days up to yesterday), the Conversion Value filter, the campaign name contains / does not contain filters (they only matched the exact name), the CpcBid vs `avg` filter, keywords that use the ad group bid, bids rounded to 0.01, unchanged bids no longer re-sent and logged, and excluded product groups are skipped. No sheet changes needed.
