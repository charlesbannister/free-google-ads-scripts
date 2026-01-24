# Advanced Anomaly Detector

Receive customised alerts via Slack and Email when anomalies are detected across your Google Ads Account, Campaigns, Labels or Ad Groups.

## Details

| | |
|---|---|
| **Category** | Monitoring |
| **Tags** | Monitoring, Analysis, Alerts, Performance, Reporting, Optimization |
| **Difficulty** | Advanced |
| **Schedule** | Daily |
| **Makes Changes** | No |
| **Last Updated** | 2024-08-01 |

## Links

- [Template Spreadsheet](https://docs.google.com/spreadsheets/d/1Fcaq3PGgBpSwTqs-PAWEo2lkP7WgyQ4NTRZjC4yuIhI)
- [YouTube Tutorial](https://www.youtube.com/watch?v=6ELoD2o3Z0o)
- [Script on GitHub](https://raw.githubusercontent.com/charlesbannister/free-google-ads-scripts/refs/heads/master/advanced_anomaly_detector/advanced_anomaly_detector.js)

## New in August 2024! 🚀🌟

## Set up alerts when:

- The Account is down (zero impressions)
- A Campaign's cost spikes
- An Ad Group's clicks dip suddenly
- A Campaign Label falls under a set number of conversions

## All including customisable:

- Date ranges
- Filters
- Alert thresholds
- Notification frequency

## Slack Notifications in 5 minutes

To enable Slack Notifications you'll need a Webhook. That's a URL Slack provide so the script knows where to send the alerts.

We've written a guide on grabbing that URL. It's dead easy, promise!

[How to create a Slack Webhook (to enable Slack Notifications)](https://docs.google.com/document/d/1g1CX6ZRMtmx6KjNTjxMutqZp5vca7ESs-I0ztl0LcEM/edit?tab=t.0#heading=h.tkx17hheq8ll)

Once you've got the URL, add it to the top of the script by editing the `SLACK_TEAM_WEBHOOK_URL` variable. See the YouTube video at 5:45 for more information.

## What if I have suggestions?

Please let me know! Hearing your pain points is the number one way I can make improvements.
