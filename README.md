Warning: fs was declared with const; use let for reassignable variables.
# GMB Sell Reviews dashboard

Static GitHub Pages dashboard built from the supplied GMB Sell Reviews and SI_DATA sheets. The current source snapshot is embedded in `index.html`, so it renders when opened directly as a local file and when published to GitHub Pages. `data.json` is retained as a separate readable copy of the snapshot.

## Views

- MTD, selected day, full month and custom dates
- MTD comparison against the previous calendar month using the same day-of-month cutoff
- Review totals, derived positive/negative counts and review-to-stock-in %
- Ranked top regions, minimum stock-in eligibility threshold, and dashboard risk signals
- Daily trend explorer and region-wise negative review visibility with search and CSV export
- Collector event view with collector ID, date, location/store, tracking-status filter, review links and Drive screenshot links

## Data snapshot

data.json contains review and stock-in source values through **17 November 2026**. collector-data.js contains event rows with a collector ID from **GMB Review Event Capture · QR Scan Events** through **18 November 2026**. To refresh, export/copy the current month region-by-day table and Negative Reviews rows from the GMB Sell Reviews workbook, refresh SI_DATA from Customer Review Form (Responses), and refresh QR Scan Events from GMB Review Event Capture. Keep only stock-in rows where DRIVER_UPLOAD_START is nonblank and VEHICLE_CONDITION exactly matches Non Scrap & Perfect Working Condition Car. Dates use the sheet's Asia/Calcutta date.

## Definitions and assumptions

- Positive = total reviews − listed negative review rows, floored at zero.
- Review / Stock-In % = reviews ÷ filtered stock-ins. It is a comparison ratio, not a conversion rate.
- MTD comparison uses matching day numbers in the selected month and previous calendar month. If the previous month is absent, it is shown as unavailable rather than zero.
- Top-region ranking includes only regions with at least 100 qualified stock-ins for the selected period.
- Risk flags are dashboard triage assumptions: at least 10 negative reviews and at least 5% negative share, or at least 100 stock-ins and a review/stock-in ratio below 15%. These are not formal business targets.
- Collector “collected” counts mean tracking_status = screenshot_uploaded. The source labels the screenshot as customer-provided evidence and does not verify Google publication. Other tracking states remain selectable.
- Negative Reviews does not populate its Region field in the sampled rows. The dashboard infers region from city/state text for regional visibility; unmatched rows are marked Unmapped. Verify this mapping before operational use.
- The source review workbook contains October and November 2026 tabs. Snapshot ends at 17 November, aligned to the latest date in those source tables and the SI_DATA extract.
- Review/stock-in data is embedded in the HTML to allow direct local opening. Collector rows load from collector-data.js; keep that file beside index.html when publishing.

## Source workbook update status

The stock-in workbook read succeeded, and the dashboard snapshot applies the requested filters and groups by RETAIL_REGION and SI_DATA column AA. A `GMB Stock-In DOD` tab was created in the source workbook, but the subsequent update call returned Google Sheets API `404 NOT_FOUND`; its contents could not be confirmed or written in this session.
