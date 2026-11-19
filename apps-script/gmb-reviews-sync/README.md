# GMB Reviews Apps Script Sync

This standalone Apps Script project syncs the existing Sell and stock-in snapshot, refreshes the Buy Small Cities aggregates from the `Buy Reviews` sheet, and polls Gmail for the partner's CARS24 MTD Excel report. Gmail polling runs every five minutes and is near real time, not an instant push. Public views receive aggregate counts only.

## Email report view

- Finds recent `.xlsx` messages using the optional Gmail search query in Script Properties (`MAIL_QUERY`). The default searches recent XLSX attachments; the attachment name must match `CARS24_Reviews_MTD_*.xlsx`.
- Converts the workbook temporarily to Google Sheets, reads `Reviews Data`, checks its row and 1–2 star counts against `Summary`, then trashes the temporary converted file.
- Publishes only Buy, Sell and Other totals grouped by review date and state/region to `mail-mtd-data.json`. It does not publish review text, reviewer names, phone numbers, or individual listings.
- Classifies a Business Name containing “Buy” as Buy; otherwise a name containing “Sell” as Sell; remaining listings stay in Other. Negative means 1–2 stars. Positive is total minus negative, so neutral is included in positive in line with the dashboard convention.
- Regions come from state names in the Location field where present, with common city names mapped to states; unresolved locations appear under Unmapped. The workbook date range/as-of is preserved; reports may lag collection by 1–2 days.

## Google Apps Script setup

1. Create/open the standalone project at [script.google.com](https://script.google.com/) and copy `Code.gs` and `appsscript.json` into it. Enable the **Gmail API** and **Drive API** under **Services** in the Apps Script editor (the manifest also declares them).
2. Add Script Properties:
   - `REVIEWS_SHEET_ID`: GMB Sell Reviews spreadsheet ID.
   - `STOCK_SHEET_ID`: Customer Review Form (Responses) spreadsheet ID.
   - `GITHUB_REPOSITORY`: `ampcars24-cmyk/Sell_GMB_reviews`.
   - `GITHUB_BRANCH`: `main`.
   - `GITHUB_FILE_PATH`: `data.json`.
   - `GITHUB_TOKEN`: fine-grained GitHub token limited to this repository with **Contents: Read and write**.
   - `MAIL_QUERY` (optional): Gmail search query to restrict sender/subject, for example `from:reports@example.com filename:xlsx newer_than:120d`. Keep the report filename pattern `CARS24_Reviews_MTD_*.xlsx`.
3. Run `installSyncTriggers` once from the Apps Script editor and approve the spreadsheet read, Gmail read, Drive file, and external request access. This installs the existing sheet sync and a five-minute Gmail poll. Apps Script runs as the account that installs the trigger.
4. Check **Executions** for errors and run `getSyncStatus` to see the latest sync timestamps. Run `removeSyncTriggers` to disable all installed sync triggers.

The GitHub token must remain in Script Properties and must never be placed in source code. Updating this repository's Apps Script files does not update a separately deployed Apps Script project by itself; copy the latest source into that project and install/authorize triggers there.

## Existing Google Sheets sync

The `syncFromSheets` handler updates `data.json` from the monthly regional review summary and qualified `SI_DATA` rows. It also reads only the required fields from the `Buy Reviews` tab and updates aggregate review counts, reply rates, negative reply coverage, and replies within 24 hours for Kochi, Kolkata, Ranchi, Dehradun, and Vizag in `small-city-metrics.json`. The Buy view assumes a 24-hour SLA threshold. Negative-review replies are a response-coverage proxy, not proof of customer recovery. Qualified stock-ins require RETAIL_REGION, date from column AA, nonblank DRIVER_UPLOAD_START, and VEHICLE_CONDITION exactly `Non Scrap & Perfect Working Condition Car`. Public negative-review output is reduced to counts and omits row-level text and identity details.

