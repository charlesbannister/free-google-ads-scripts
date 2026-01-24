# Auto Negative Keywords Script (Google Sheets)

Auto-add negative keywords and receive irrelevant search term alerts without chaining yourself to search term reports.

## Details

| | |
|---|---|
| **Category** | Negative Keywords |
| **Tags** | Negative Keywords, Google Shopping, Dynamic Search Ads, Google Search |
| **Difficulty** | Easy |
| **Schedule** | Hourly |
| **Makes Changes** | Yes |
| **Last Updated** | 2026-01-23 |
| **⭐ Superstar** | Yes |

## Links

- [YouTube Tutorial](https://www.youtube.com/watch?v=-r_sklAZ8Kg)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/auto_negative_keywords_sheets/auto_negative_keywords_sheets.js)

## Perfect For

- Shopping campaigns
- Dynamic search ads (DSAs)
- Broad match keywords
- Text Ads (combat Google's "close-match" going awry)

## Once Irrelevant Search Terms Are Found You Can

- Auto-add them as negative keywords to Ad Groups or Lists
- Review them in the sheet
- Receive an alert email (one per day per sheet to avoid alert overload!)

**It boils down to this: You know your products/services better than Google.**

Use your product or service knowledge to setup "positive keywords" and cut wasted spend on irrelevant terms.

## Features

After years of improvements and hundreds of hours of development, it's jam packed with features:

- "Pick up where it left off" support
- Safety rails
- Error and warning emails (you'll want to know if this stops working)
- Multiple Ad Groups, Multiple Campaigns (sheets)
- Option to only add new negative keywords
- Negative keyword list support
- View the negative keywords in the sheet
- Advanced error logging
- Advanced query filters
- Regex Match 🆕
- Detailed email alerts 🆕
- Preview Mode 🆕
- Approx Match 🆕
- Campaign Level Negative Keywords 🆕
- Output Negatives to Sheet 🆕

## How It Works

You define a list of positive keywords in a Google Spreadsheet, the script then adds negative keywords where the query doesn't match any of the words provided.

Let's imagine you sell rucksacks and holdalls. Based on historical performance, you know queries containing "rucksack", "holdall" or "bag" are all profitable, other queries just waste money.

You would just need to add "ruc", "hold" and "bag" to the Google Sheet and where queries don't contain "ruc", "hold" or "bag" they will be automatically added as a negative keyword.

This can be setup for dozens of words, through dozens of adGroups in dozens of campaigns. Just create a new tab for each campaign. The tab (or sheet) can be named anything.

The negative keywords are added at AdGroup level or to Negative Keyword Lists.

**Script timing out?** No problem, next time the script runs it will start with any rules it didn't get round to last time.

## Tips for Best Results

This script has been used across thousands of accounts and we also use it ourselves. It works a charm when setup correctly. Here are some important tips:

1. **Set the script to run hourly** so it spots irrelevant search terms as soon as they appear
2. **Set an impression minimum** for larger accounts to ensure too many negative keywords aren't added (there's a 20k limit at Ad Group level)
3. **Partial match is still a match** - If the positive keyword is "ruc" and the term is "rucsac" it will be allowed through (it will not be added as a negative keyword)
4. **Case insensitive** - You don't need to add multiple versions for different cases. Never repeat words especially if the number of matches is more than one.
5. **Always preview first** - Review the results before running
6. **Enable Alert Mode** so you can keep an eye on what's being added
