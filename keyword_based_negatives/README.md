# Keyword-Based Negatives Script

**Authors:** Charles Bannister & Gabriele Benedetti  
**Version:** 1.0.0

A Google Ads Script that automatically identifies potential negative keywords by comparing search terms against your account's actual keywords within the same ad group.

## How It Works

1. **Scans Ad Groups** - Loops through all ad groups matching your campaign/ad group name filters
2. **Compares Search Terms to Keywords** - For each ad group, compares search terms against the keywords using fuzzy matching
3. **Identifies Mismatches** - Flags search terms that don't match any keyword in their ad group
4. **Outputs to Sheet** - Writes potential negatives to a Google Sheet with checkboxes for review
5. **Applies on Next Run** - On subsequent runs, checked items are applied as negative keywords

## Features

- **Fuzzy Matching** - Uses `SearchTermMatcher` with configurable similarity threshold (default 80%)
- **Performance Filters** - Filter by minimum clicks, impressions, and maximum conversions
- **Campaign/Ad Group Filters** - Include or exclude by name patterns
- **Two Modes**:
  - **Manual Review Mode** - Review and check items before applying
  - **Full Automate Mode** - Automatically apply all found negatives
- **Email Notifications** - Get notified when negatives are found or applied
- **Logging** - All applied negatives are logged with timestamp and execution mode

## Setup

1. **Create a Google Sheet** - Type `sheets.new` in your browser
2. **Copy the Sheet URL** - Paste it into the `SPREADSHEET_URL` variable
3. **Configure Settings** - Adjust the configuration variables as needed
4. **Run the Script** - Preview first to verify results

## Configuration Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `SPREADSHEET_URL` | - | Your Google Sheet URL |
| `OUTPUT_SHEET_NAME` | "Output" | Sheet for potential negatives |
| `LOGS_SHEET_NAME` | "Logs" | Sheet for applied negatives log |
| `LOOKBACK_DAYS` | 30 | Days of search term data to analyze |
| `FULL_AUTOMATE_MODE` | false | Auto-apply all found negatives |
| `DEBUG_MODE` | false | Enable verbose logging |
| `MIN_CLICKS` | 0 | Minimum clicks threshold |
| `MIN_IMPRESSIONS` | 1 | Minimum impressions threshold |
| `MAX_CONVERSIONS` | 0 | Maximum conversions (0 = non-converting only) |
| `FUZZY_MATCH_THRESHOLD` | 80 | Similarity threshold (0-100) |
| `CAMPAIGN_NAME_CONTAINS` | [] | Include campaigns containing these strings |
| `CAMPAIGN_NAME_NOT_CONTAINS` | [] | Exclude campaigns containing these strings |
| `AD_GROUP_NAME_CONTAINS` | [] | Include ad groups containing these strings |
| `AD_GROUP_NAME_NOT_CONTAINS` | [] | Exclude ad groups containing these strings |
| `EMAIL_RECIPIENTS` | "" | Comma-separated email addresses |
| `NEGATIVE_MATCH_TYPE` | "EXACT" | Match type for negatives (EXACT/PHRASE/BROAD) |

## Sheet Structure

### Output Sheet
Contains potential negatives with checkboxes:
- Check the box next to items you want to negate
- On next run, checked items will be applied
- Sheet is cleared after processing

### Logs Sheet
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
4. Run script again → checked items are applied
5. Repeat

### Full Automate Mode
1. Script finds potential negatives
2. Automatically checks all boxes
3. Immediately applies all as negatives
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

## Version History

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
