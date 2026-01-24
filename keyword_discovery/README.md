# Keyword Discovery Script (with Fuzzy Match)

Surface successful search terms which haven't yet been added as keywords. Great for finding new Search opportunities with built-in fuzzy matching.

## Details

| | |
|---|---|
| **Category** | Keywords |
| **Tags** | Keywords, Search Terms, Google Search, Analysis, Reporting |
| **Difficulty** | Intermediate |
| **Schedule** | Weekly |
| **Makes Changes** | No |
| **Last Updated** | 2026-01-20 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1yCg93xb3Yhx1c8q9AIbA9KHfa5cyMIquzvA5GW3OOH4)
- [YouTube Tutorial](https://www.youtube.com/watch?v=Pj0rflybGRM)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/keyword_discovery/keyword_discovery.js)

## Overview

This script will surface successful search terms which haven't yet been added as keywords. It's great for finding new Search opportunities.

## How It Works

There are two main steps:

1. Grab search terms based on your filters (these generally want to be successful search terms)
2. Populate the sheet (and optionally send an alert) if any of the search terms are NOT in as keywords

The script can be run manually as part of your internal processes or on a schedule. Enable notifications to receive an email when there are new opportunities.

Using the advanced filters, you can define success however you like. That includes filtering search terms based on text they contain (or don't contain).

## Introducing Fuzzy Matching

The first version of this script was producing a lot of accurate but unhelpful results.

If "purple dog collars" has been added as a keyword then "purple dog collar" isn't much of an opportunity thanks to close variant matching.

That's where fuzzy matching comes in. Similar to close variant matching, the script won't report opportunities (or alert you) if there's an existing, similar keyword.

### Fuzzy Match Examples

| Search Term | Keyword | Fuzzy Match Score |
|-------------|---------|-------------------|
| dog collars | dog collar | 91% |
| purple dog collars | dog collars | 61% |
| purple dog collars | purple dog collar | 94% |
| purple dog collars | pink dog collars | 72% |
| cat collars | dog collar | 63% |
| stories for dogs | story for dogs | 81% |

## Settings

Settings depend on goals but it's generally wise to setup at least two rules:

- **Successful search terms**: high clicks, with conversions at a good ROI
- **High clicks**: if they're relevant, adding search terms as keywords can help keep an eye on performance and reveal Quality Score metrics

## Adding New Rules

Adding new rules is simple:

1. Copy and paste an existing row
2. Duplicate an existing output sheet
3. Name the new sheet
4. Enter the same name in the Output Sheet Name column so the script knows where to store opportunities

## What if I have suggestions?

Please let me know! Hearing your pain points is the number one way I can make improvements.
