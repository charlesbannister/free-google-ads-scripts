# Performance Max Placement Exclusions

This script helps you manage placement exclusions for Performance Max and Display campaigns. It pulls placement data, writes it to a Google Sheet with checkboxes, and adds selected placements to a shared exclusion list. Optionally, it can use ChatGPT to analyze website content and help identify irrelevant placements.

## Quick Start

1. **Copy the template sheet**: [Click here to make a copy](https://docs.google.com/spreadsheets/d/18vdejatcc7b3cWLtNGdmpVFoSDQ4JXXvDiVDLRpAUb4/copy)
2. **Create a shared exclusion list** in Google Ads: Tools & Settings > Shared Library > Placement exclusions
3. **Add the script** in Google Ads: Bulk actions > Scripts > New script
4. **Paste the script** and update `SPREADSHEET_URL` with your copied sheet URL
5. **Configure settings** in your sheet's Settings tab
6. **Preview the script** to test it works

## How It Works

1. **Fetches placement data** - Queries Performance Max and Display campaigns for placement information (websites, YouTube videos, mobile apps, Google Products)
2. **Writes to output sheets** - Creates four output sheets, one per placement type, with checkboxes in the first column
3. **Optional ChatGPT analysis** - For website placements, ChatGPT can analyze page content to help identify irrelevant sites
4. **You select what to exclude** - Check the boxes next to placements you want to exclude
5. **Adds to exclusion list** - On next run, adds all checked placements to your shared exclusion list

## Sheets

### Settings Sheet

Main configuration options:

| Setting                    | Description                                                      |
| -------------------------- | ---------------------------------------------------------------- |
| Shared Exclusion List Name | Name of your placement exclusion list in Google Ads (must exist) |
| Lookback Window (Days)     | How many days back to look for placement data                    |
| Minimum Impressions        | Only show placements with at least this many impressions         |
| Minimum Clicks             | Only show placements with at least this many clicks              |
| Minimum Cost               | Only show placements with at least this cost                     |
| Maximum Conversions        | Exclude placements with more than this many conversions          |
| Max Results                | Limit the number of results (0 = no limit)                       |
| Campaign Name Contains     | Filter to campaigns containing this text                         |
| Campaign Name Not Contains | Exclude campaigns containing this text                           |
| Enabled Campaigns Only     | Checkbox to only include enabled campaigns                       |

**Placement Type Filters** - Enable/disable each placement type (YouTube, Website, Mobile App, Google Products) and optionally enable automation per type.

### Settings: Placement Filters Sheet

Filter placements using text matching lists. Each column has an enable checkbox at the top:

- **Placement Contains / Not Contains** - Filter by placement ID/URL
- **Display Name Contains / Not Contains** - Filter by display name
- **Target URL Contains / Not Contains** - Filter by target URL
- **Target URL Ends With / Not Ends With** - Filter by URL suffix (useful for TLDs like `.ru`, `.cn`)

### Settings: ChatGPT Sheet

Optional AI-powered website analysis:

| Setting               | Description                                                        |
| --------------------- | ------------------------------------------------------------------ |
| Enable ChatGPT        | Turn ChatGPT integration on/off                                    |
| API Key               | Your OpenAI API key                                                |
| Use Cached Responses  | Reuse previous ChatGPT responses for the same URLs                 |
| Prompt                | Custom prompt for ChatGPT to analyze website content               |
| Response Contains     | Flag placements where ChatGPT response contains these terms        |
| Response Not Contains | Flag placements where ChatGPT response doesn't contain these terms |

### Output Sheets

Four output sheets, one per placement type:

| Sheet                      | ChatGPT Column | Notes                                             |
| -------------------------- | -------------- | ------------------------------------------------- |
| Output: Website            | ✅ Yes         | Website placements                                |
| Output: YouTube            | ✅ Yes         | YouTube video placements                          |
| Output: Mobile Application | ❌ No          | Mobile app placements (app IDs can't be analyzed) |
| Output: Google Products    | ❌ No          | Google Products placements                        |

**Tab colors**:

- 🟢 **Green** - Sheet has placement results
- 🔴 **Red** - No results found for this placement type

### LLM Responses Cache Sheet

Stores cached ChatGPT responses to avoid repeated API calls for the same URLs.

## Script Configuration

Variables at the top of the script:

| Variable               | Description                                                            |
| ---------------------- | ---------------------------------------------------------------------- |
| `SPREADSHEET_URL`      | Your Google Sheet URL                                                  |
| `DEBUG_MODE`           | Set to `true` for detailed logging                                     |
| `ONLY_PROCESS_CHANGES` | Set to `true` to skip fetching new data and only process checked boxes |

## Important Notes

- **Metrics limitation**: The `performance_max_placement_view` resource only supports impressions. Clicks, cost, and conversions will show as 0 for PMax placements (Display placements have full metrics).
- **Google domains**: Google-owned domains (youtube.com, gmail.com, etc.) cannot be excluded due to policy and are automatically filtered out.
- **Shared exclusion list**: Must be created in Google Ads before running the script.
- **Mobile app exclusions**: Mobile apps require bulk upload to exclude and will be noted in the sheet.
- **ChatGPT API costs**: Using ChatGPT will incur OpenAI API costs based on your usage.

## Version

Current version: 2.3.0
