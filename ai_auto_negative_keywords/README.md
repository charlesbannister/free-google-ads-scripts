# AI Auto Negative Keywords

autoneg.shabba.io

## Description

Google Ads is becoming less about what you want to target, and more about what you need to exclude. With this advanced tool, you can setup intelligent rules to ensure wasteful, irrelevant search terms are excluded before money is wasted on clicks that won't convert.

This script supports Dynamic Search Ads (DSAs), Shopping Ads, Performance Max, and other campaign types, bringing back the precision control that keywords once provided while maintaining Google's intelligent matching for legitimate variations like misspellings and typos.

### Key Features

- **Custom words and keyword protection**: Check search terms against your own words and protect terms close to your campaign or ad group keywords
- **Approx. Match Setting**: Intelligent matching that won't exclude legitimate variations (e.g., "banksie" when targeting "banksy")
- **Search Term Alerts**: Get notified when performance metrics exceed thresholds
- **AI/LLM Integration**: Let AI review search terms and suggest broad match negatives or relevance scoring
- **Comprehensive Reporting**: Track results across different lookback windows with detailed KPIs
- **Perfect for Brand Campaigns**: Stricter matching control for brand protection

## Version 17

Performance Max search terms can now be negated. The script reads real Performance Max search terms (with cost) from Google's Performance Max search terms report, applies the rule's text, AI and performance conditions to them, and, when auto-apply is on, adds matches as exact negatives:

- **Campaign rules:** to the Performance Max campaign itself, or to the selected negative keyword list. A list only blocks terms in the campaigns it is applied to, so apply the list to your Performance Max campaigns.
- **Account rules:** to the selected negative keyword list.

If Google's search terms report is not available, the script falls back to the search term categories used in version 16. Categories are reported only and never negated. Set `PMAX_AUTO_APPLY = false` at the top of the script to keep all Performance Max terms report-only.

## Version 16

Use **Protect keywords** on a campaign or ad-group rule to keep search terms close to an enabled keyword. You can use it alongside **Custom words**. The script reads keywords once per account run. It skips terms it cannot compare, including Performance Max terms, and keeps scoped keyword protection report-only when a shared negative list is selected.

## Version 15

If the rules service is unavailable, the script stops before opening the spreadsheet. If a Google Ads report fails, the rule keeps its last report and applies no negatives. Missing example terms no longer interrupt the live search terms report, and Performance Max terms appear in reports with auto-apply enabled.

## Resources

- [Current standalone script](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/ai_auto_negative_keywords/ai_auto_negative_keywords.js)
- **Google Sheet Template**: https://docs.google.com/spreadsheets/d/1xwdogT2HiHV_Gx9pdY4rErhLSm7Qrooj6wbikEjf-G0/edit?gid=0#gid=0
- **Setup Interface**: https://autoneg.shabba.io
