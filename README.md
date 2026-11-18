# GMB Sell Reviews dashboard

Static GitHub Pages dashboard built from the supplied GMB Sell Reviews and SI_DATA sheets. The current source snapshot is embedded in `index.html`, so it renders when opened directly as a local file and when published to GitHub Pages. `data.json` is retained as a separate readable copy of the snapshot.

## Views

- MTD, selected day, full month and custom dates
- Review totals, derived positive/negative counts and review-to-stock-in %
- Daily review and qualified stock-in trend
- Region comparison and negative review visibility with search and CSV export

## Data snapshot

`data.json` contains source values through **17 November 2026**. To refresh, export/copy the current month region-by-day table and Negative Reviews rows from the GMB Sell Reviews workbook, then refresh SI_DATA columns F (RETAIL_REGION), N (DRIVER_UPLOAD_START), W (VEHICLE_CONDITION) and AA (date) from Customer Review Form (Responses). Keep only stock-in rows where DRIVER_UPLOAD_START is nonblank and VEHICLE_CONDITION exactly matches `Non Scrap & Perfect Working Condition Car`. Dates use the sheet's Asia/Calcutta date. Replace the corresponding sections in `data.json`.

## Definitions and assumptions

- Positive = total reviews − listed negative review rows, floored at zero.
- Review / Stock-In % = reviews ÷ filtered stock-ins. It is a comparison ratio, not a conversion rate.
- Negative Reviews does not populate its Region field in the sampled rows. The dashboard infers region from city/state text for regional visibility; unmatched rows are marked Unmapped. Verify this mapping before operational use.
- The source review workbook contains October and November 2026 tabs. Snapshot ends at 17 November, aligned to the latest date in those source tables and the SI_DATA extract.
- The HTML embeds the snapshot to avoid browsers blocking a local `fetch('./data.json')` request. When refreshing data, update both the embedded JSON block and `data.json`.

## Source workbook update status

The stock-in workbook read succeeded, and the dashboard snapshot applies the requested filters and groups by RETAIL_REGION and SI_DATA column AA. A `GMB Stock-In DOD` tab was created in the source workbook, but the subsequent update call returned Google Sheets API `404 NOT_FOUND`; its contents could not be confirmed or written in this session.
