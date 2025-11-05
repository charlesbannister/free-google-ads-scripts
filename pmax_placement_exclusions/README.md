# Performance Max Placement Exclusions

This script helps you manage placement exclusions for your Performance Max campaigns. It pulls placement data from your PMax campaigns, writes it to a Google Sheet with checkboxes, and then adds any placements you've selected to a shared exclusion list.

## How It Works

1. **Gets placement data** - The script queries your Performance Max campaigns and pulls placement information (websites, apps, etc.) where your ads appeared.
2. **Writes to sheet** - All placements are written to a "Placements" sheet with checkboxes in the first column. You can see campaign names, placement URLs, impressions, and other details.
3. **You select what to exclude** - Check the boxes next to placements you want to exclude from your campaigns.
4. **Adds to exclusion list** - The next time the script runs, it adds all checked placements to your shared exclusion list and links that list to all your PMax campaigns.

## Settings

The script uses a "Settings" sheet in your Google Sheet with the following options:

- **Shared Exclusion List Name** - The name of the shared placement exclusion list in your Google Ads account (must exist before running)
- **Lookback Window (Days)** - How many days back to look for placement data (default: 30)
- **Minimum Impressions** - Only show placements with at least this many impressions (default: 0)
- **Minimum Clicks** - Only show placements with at least this many clicks (default: 0)
- **Campaign Name Contains** - Filter to only campaigns with names containing this text (leave empty for all)
- **Campaign Name Not Contains** - Exclude campaigns with names containing this text (leave empty for none)
- **Enabled campaigns only** - Checkbox to include only enabled campaigns, or uncheck to also include paused campaigns (removed campaigns are always excluded)

## Important Notes

- The `performance_max_placement_view` resource only supports the impressions metric. Clicks, cost, conversions, and other metrics will show as 0 in the sheet.
- Google-owned domains (like youtube.com, gmail.com) cannot be excluded and will be automatically filtered out.
- The shared exclusion list must be created in Google Ads before running the script (Tools & Settings > Shared Library > Placement exclusions).
