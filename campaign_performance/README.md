# 📊 Weekly/Monthly Campaign Performance Report Script

A comprehensive Google Ads script that generates automated weekly campaign performance reports with color-coded metrics, target tracking, and email notifications.

## 🎯 What This Script Does

This script creates detailed performance reports comparing campaign metrics across multiple time periods, helping you identify trends and track performance against weekly targets.

### Key Features

- **Multi-Period Comparison**: Compare campaign performance across multiple time periods (default: 3 periods of either weeks or months)
- **Customizable Metrics**: Choose which metrics to track via a settings sheet
- **Target Tracking**: Set weekly targets and get visual color-coding based on performance
- **Automated Email Reports**: Beautiful HTML email reports sent to specified recipients
- **MCC Support**: Can run across multiple accounts from a Manager (MCC) account
- **Google Sheets Integration**: Automatically writes data to a Google Sheets report

## 📈 Supported Metrics

- **Volume Metrics**: Impressions, Clicks, Cost, Conversions
- **Rate Metrics**: CTR (%), Conversion Rate (%), CPC, Cost/Conversion
- **Search Impression Share Metrics**: Search Impr. Share (%), Search Top IS (%), Search Abs. Top IS (%)

## 🔧 Setup Instructions

### 1. Create a Google Sheets Report

1. Go to [sheets.new](https://sheets.new) to create a new spreadsheet
2. Copy the spreadsheet URL from your browser
3. Update the `SPREADSHEET_URL` constant in the script with your URL

### 2. Configure Account IDs (for MCC accounts)

If running from a Manager account:

- Update the `ACCOUNT_IDS` array with the account IDs you want to report on
- Format: `['123-456-7890', '987-654-3210']`

### 3. Install and Run

1. Copy the script into Google Ads Scripts
2. Run the script once to create the settings sheet structure
3. Configure your preferences in the generated settings sheet

## ⚙️ Configuration Options

The script automatically creates a `settings` sheet with the following options:

### Report Configuration

- **Weeks in Period**: Number of weeks per time period (default: 4)
- **Number of Periods**: How many time periods to compare (default: 3)
- **Email Addresses**: Comma-separated emails for report delivery (optional)

### Metric Selection

Enable/disable any combination of these metrics:

- ✅ Impressions
- ✅ Clicks
- ✅ Cost
- ✅ CTR (%)
- ✅ CPC
- ✅ Conversions
- ✅ Conv. Rate (%)
- ✅ Cost/Conv.
- ✅ Search Impr. Share (%)
- ✅ Search Top IS (%)
- ✅ Search Abs. Top IS (%)

## 🎨 Color-Coded Performance

The script provides visual performance indicators:

- 🟢 **Green**: Performance meets or exceeds targets
- 🟡 **Yellow**: Performance is close to targets (80-100% for volume metrics)
- 🔴 **Red**: Performance is below targets

### Target Logic

- **Volume Metrics** (impressions, clicks, cost, conversions): Weekly targets are multiplied by the number of weeks in the period
- **Rate Metrics** (CTR, conversion rate, etc.): Weekly targets remain the same for any time period

## 📧 Email Reports

When email addresses are configured, the script sends beautiful HTML email reports featuring:

- 📊 Executive summary with key statistics
- 📈 Detailed performance table with color-coding
- 💰 Budget and bidding strategy information
- 📅 Clear date ranges and time period labels

## 🔄 How It Works

1. **Data Collection**: Retrieves campaign data for each configured time period using Google Ads API
2. **Processing**: Calculates derived metrics (CTR, CPC, conversion rate, etc.)
3. **Target Comparison**: Compares actual performance against weekly targets
4. **Visualization**: Applies color-coding based on performance vs targets
5. **Reporting**: Writes data to Google Sheets and optionally sends email reports

## 📊 Sample Output Structure

```
Campaign Name: Brand Campaign
Bidding Strategy: TARGET_CPA - Smart Bidding Strategy
Budget: $1,500/month (est.)

Metric          | Weekly Target | Period 1 | Period 2 | Period 3
----------------|---------------|----------|----------|----------
Impressions     | 10,000       | 42,500   | 38,200   | 35,800
Clicks          | 500          | 2,150    | 1,920    | 1,780
Cost            | $200         | $825     | $745     | $695
CTR (%)         | 5.00%        | 5.06%    | 5.03%    | 4.97%
```

## 🛠️ Technical Details

- **API**: Uses Google Ads API v20 with GAQL queries
- **Data Processing**: Handles multiple time periods with proper date range calculations
- **Error Handling**: Comprehensive error handling with helpful debugging information
- **Performance**: Optimized queries and data processing for large account structures

## 🔍 Debug Mode

Set `DEBUG_MODE = true` to see detailed logging including:

- GAQL query validation links
- Row counts and sample data
- Processing steps and timing information

## 📋 Requirements

- Google Ads account with script access
- Google Sheets for report storage
- Valid campaign data in the specified time periods
- (Optional) Email addresses for automated report delivery

## 🆚 Version History

- **v3.2.0**: Current version with full multi-period support and email reporting
- Enhanced target tracking and color-coding system
- Improved email HTML formatting and styling

---

_Generated by Google Ads Performance Script - Helping you track and optimize campaign performance across time periods_
