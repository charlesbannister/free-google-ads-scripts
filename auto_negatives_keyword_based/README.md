# Auto Negative Keywords Script (Keyword Based)

**Authors:** Charles Bannister & Gabriele Benedetti  
**Version:** 1.4.0

A Google Ads Script that automatically identifies potential negative keywords by comparing search terms against your account's actual keywords within the same ad group.

## How It Works

1. **Scans Ad Groups** - Loops through all ad groups matching your campaign/ad group name filters
2. **Compares Search Terms to Keywords** - For each ad group, compares search terms against the keywords using fuzzy matching
3. **Identifies Mismatches** - Flags search terms that don't match any keyword in their ad group (below 80% similarity)
4. **Outputs to Sheet** - Writes potential negatives to a Google Sheet with checkboxes for review
5. **Applies on Next Run** - On subsequent runs, checked items are applied as exact match negative keywords

## Features

- **Fuzzy Matching** - Uses `SearchTermMatcher` with configurable similarity threshold (default 80%)
- **Multiple Match Methods** - Checks whole term, spaceless, sorted words, sliding window, subset words, and shared words
- **Performance Filters** - Filter by minimum clicks, impressions, and maximum conversions
- **Campaign/Ad Group Filters** - Include or exclude by name patterns
- **Two Modes**:
  - **Manual Review Mode** - Review and check items before applying
  - **Full Automate Mode** - Automatically apply all found negatives
- **Email Notifications** - Get notified when negatives are found or applied
- **Logging** - All applied negatives are logged with timestamp and execution mode
- **Built-in Test Suite** - Verify matching logic before running on live data

## Setup

1. **Create a Copy of the Template** - [Click here to copy the template](https://docs.google.com/spreadsheets/d/1x3XcN3EwJzo4RzDXphtSINjC5yKIhCd81JJ00CLwQT0/copy)
2. **Copy Your Sheet URL** - Paste it into the `SPREADSHEET_URL` variable in the script
3. **Configure Settings** - Adjust the configuration variables as needed
4. **Preview the Script** - Always preview first to verify results before running live

## Configuration Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `SPREADSHEET_URL` | - | Your Google Sheet URL (copy from template) |
| `OUTPUT_SHEET_NAME` | "Output" | Sheet for potential negatives |
| `LOGS_SHEET_NAME` | "Logs" | Sheet for applied negatives log |
| `LOOKBACK_DAYS` | 30 | Days of search term data to analyze |
| `FULL_AUTOMATE_MODE` | false | Auto-apply all found negatives |
| `DEBUG_MODE` | false | Enable verbose logging |
| `RUN_TEST_SUITE` | false | Run matching logic tests before script |
| `MIN_CLICKS` | 5 | Minimum clicks threshold |
| `MIN_IMPRESSIONS` | 1 | Minimum impressions threshold |
| `MAX_CONVERSIONS` | 0 | Maximum conversions (0 = non-converting only) |
| `FUZZY_MATCH_THRESHOLD` | 80 | Similarity threshold (0-100) |
| `CAMPAIGN_NAME_CONTAINS` | [] | Include campaigns containing these strings |
| `CAMPAIGN_NAME_NOT_CONTAINS` | [] | Exclude campaigns containing these strings |
| `AD_GROUP_NAME_CONTAINS` | [] | Include ad groups containing these strings |
| `AD_GROUP_NAME_NOT_CONTAINS` | [] | Exclude ad groups containing these strings |
| `EMAIL_RECIPIENTS` | "" | Comma-separated email addresses |
| `NEGATIVE_MATCH_TYPE` | "EXACT" | Match type for negatives (EXACT/PHRASE/BROAD) |

## Output Sheet Structure

| Column | Description |
|--------|-------------|
| Negate | Checkbox - check to add as exact match negative |
| Search Term | The search term that triggered ads |
| Keywords | Ad group keywords (first 10 by clicks) |
| Match Score | Highest similarity score found (%) |
| Campaign Name | Parent campaign |
| Ad Group Name | Parent ad group |
| Campaign ID | Campaign identifier |
| Ad Group ID | Ad group identifier |
| Impressions | Total impressions |
| Clicks | Total clicks |
| Cost | Total cost |
| Conversions | Total conversions |
| Conv. Value | Total conversion value |
| CTR | Click-through rate |
| Conv. Rate | Conversion rate |
| CPA | Cost per acquisition |
| ROAS | Return on ad spend |
| Date Found | Date the term was flagged |

## Logs Sheet

Append-only log of applied negatives:
- Timestamp
- Search term, campaign, ad group details
- Execution mode (Live/Preview)
- Status (Added/Failed)

## Workflow

### Manual Review Mode (Recommended)
1. Script finds potential negatives → writes to Output sheet
2. Review the Output sheet
3. Check boxes for items you want to negate
4. Run script again → checked items are applied as exact match negatives
5. Repeat

### Full Automate Mode
1. Script finds potential negatives
2. Automatically checks all boxes
3. Immediately applies all as exact match negatives
4. ⚠️ Use with caution - no manual review!

## Example Filters

```javascript
// Only run on Brand campaigns
const CAMPAIGN_NAME_CONTAINS = ['Brand'];

// Exclude DSA and PMAX campaigns
const CAMPAIGN_NAME_NOT_CONTAINS = ['DSA', 'PMAX', 'Performance Max'];

// Only run on Exact and Phrase match ad groups
const AD_GROUP_NAME_CONTAINS = ['Exact', 'Phrase'];
```

## Matching Logic

The script uses multiple matching methods to determine if a search term is relevant:

1. **Whole term match** - Direct fuzzy comparison
2. **Spaceless match** - Ignores spaces ("runningshoes" = "running shoes")
3. **Sorted word match** - Ignores word order ("shoe running" = "running shoe")
4. **Sliding window** - Finds keyword within longer search term
5. **Subset words** - All search term words exist in keyword
6. **Shared words** - Any keyword word matches any search term word

If any method scores ≥80% (configurable), the search term is considered a match and NOT flagged.

## Version History

- **1.4.0** - Enhanced matching logic & output improvements
  - Added subset words and shared words matching
  - Match Score column moved after Keywords
  - Added "Negate" header with explanatory note
  - Built-in test suite with RUN_TEST_SUITE option
  - Output sorted by clicks (highest first)
  - Column formatting (integers, decimals, percentages)

- **1.3.0** - Matching logic improvements
  - Added sorted word matching (word order independent)
  - Added sliding window matching
  - Improved score consistency between decision and display

- **1.2.0** - Added spaceless matching
  - "runningshoes" now matches "running shoes"

- **1.1.0** - Added keywords column
  - Keywords displayed next to search terms (first 10, ordered by clicks)
  - Shows ellipsis if more than 10 keywords

- **1.0.0** - Initial release
  - Fuzzy matching via SearchTermMatcher
  - Checkbox-based review workflow
  - Full automate mode option
  - Email notifications
  - Comprehensive logging

## Support

For questions, feedback, or customizations, contact Charles Bannister:
- [LinkedIn](https://www.linkedin.com/in/charles-bannister/)
