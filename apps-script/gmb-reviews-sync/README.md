# GMB Reviews Apps Script Sync

This Apps Script project syncs the Sell review summary and filtered SI_DATA rows to the existing GitHub Pages snapshot at `data.json`. It is an automation project, not an HTML dashboard.

## What it syncs

- Monthly review counts by region and date from tabs named like `Oct'26` and `Nov'26`.
- Qualified SI_DATA counts grouped by RETAIL_REGION and column AA date, requiring nonblank DRIVER_UPLOAD_START and VEHICLE_CONDITION exactly equal to `Non Scrap & Perfect Working Condition Car`.
- Negative-review counts by date and inferred region. Review text, reviewer identity, business location, and other row-level details are omitted from the public sync.

## Near-real-time behavior

Installable on-edit, on-change, and form-submit triggers start a sync when either source spreadsheet changes. A one-minute time trigger is a fallback for formula/API-driven changes that do not fire edit triggers. GitHub Pages must then finish its normal deployment; end-to-end updates can take longer than a minute. Apps Script triggers are best-effort and are subject to Google quotas.

## Install

1. Create a standalone project at [script.google.com](https://script.google.com/) and add `Code.gs` and `appsscript.json` from this folder.
2. In **Project Settings → Script Properties**, add:
   - `REVIEWS_SHEET_ID`: the GMB Sell Reviews spreadsheet ID.
   - `STOCK_SHEET_ID`: the Customer Review Form (Responses) spreadsheet ID.
   - `GITHUB_REPOSITORY`: `ampcars24-cmyk/Sell_GMB_reviews`.
   - `GITHUB_BRANCH`: `main`.
   - `GITHUB_FILE_PATH`: `data.json`.
   - `GITHUB_TOKEN`: a fine-grained GitHub token restricted to this repo with **Contents: Read and write**.
3. Run `installSyncTriggers` once and approve the requested spreadsheet-read and external-request scopes.
4. Check **Executions** in Apps Script after the first sync. Run `getSyncStatus` to see the last successful push and installed trigger count.

Keep the GitHub token only in Script Properties. Do not paste it into source code or commit it to the repository.

## Disable

Run `removeSyncTriggers` to stop the scheduled and sheet-change syncs.

